"""Wspolna logika kontrolerow COM.

Kontroler to cienka warstwa nad obiektem ``Application`` jednej aplikacji
Office. Metody oznaczone dekoratorem :func:`action` staja sie akcjami
protokolu Bridge - reszta klasy to zwykle helpery.

Kazda akcja wykonuje sie w watku COM swojej aplikacji (patrz
:mod:`bridge.connection_manager`), a wszystkie wyjatki - w tym surowe
``pywintypes.com_error`` - sa tlumaczone na hierarchie z
:mod:`bridge.utils.errors`.
"""

from __future__ import annotations

import contextlib
import inspect
import logging
import os
from typing import Any, Callable, Iterator, TypeVar

from bridge.connection_manager import AppConnection
from bridge.utils.com_helpers import com_error, normalize_path, to_python
from bridge.utils.errors import (
    BridgeError,
    ComConnectionError,
    DocumentNotFoundError,
    InvalidReferenceError,
    ProtocolError,
    UnsupportedOperationError,
)

logger = logging.getLogger(__name__)

F = TypeVar("F", bound=Callable[..., Any])

ACTION_ATTR = "_bridge_action"


def action(name: str | None = None) -> Callable[[F], F]:
    """Oznacza metode kontrolera jako akcje dostepna przez protokol Bridge."""

    def decorator(func: F) -> F:
        setattr(func, ACTION_ATTR, name or func.__name__)
        return func

    return decorator


_DISCONNECTED = {
    -2147417848,
    -2147023174,
    -2147023170,
    -2146959355,
    -2147221021,
}

_BUSY = {
    -2147418111,  # RPC_E_CALL_REJECTED - apka zajeta (np. otwarty dialog)
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
}

_BAD_INDEX = {
    -2147352565,  # DISP_E_BADINDEX
    -2147352570,  # DISP_E_UNKNOWNNAME
}

_NOT_SUPPORTED = {
    -2147352573,  # DISP_E_MEMBERNOTFOUND
    -2147467263,  # E_NOTIMPL
}

_FILE_ERRORS = {
    -2147024894,  # ERROR_FILE_NOT_FOUND
    -2147024893,  # ERROR_PATH_NOT_FOUND
}


class BaseController:
    """Baza kontrolerow - routing akcji, mapowanie bledow, wspolne helpery."""

    APP_KEY: str = ""
    DISPLAY_NAME: str = ""
    ALERTS_OFF: Any = False

    def __init__(self, connection: AppConnection) -> None:
        self.connection = connection

    @classmethod
    def actions(cls) -> dict[str, Callable[..., Any]]:
        """Zwraca mape ``nazwa akcji -> metoda`` zbudowana z dekoratorow."""
        cached = cls.__dict__.get("_actions_cache")
        if cached is not None:
            return cached

        found: dict[str, Callable[..., Any]] = {}
        for klass in reversed(cls.__mro__):
            for member in vars(klass).values():
                action_name = getattr(member, ACTION_ATTR, None)
                if action_name:
                    found[action_name] = member

        cls._actions_cache = found  # type: ignore[attr-defined]
        return found

    @property
    def app(self) -> Any:
        """Obiekt ``Application`` - laczy sie leniwie przy pierwszym uzyciu."""
        return self.connection.application()

    def dispatch(self, action_name: str, params: dict[str, Any]) -> Any:
        """Wykonuje akcje w watku COM aplikacji i zwraca wynik gotowy do JSON-a."""
        handler = self.actions().get(action_name)
        if handler is None:
            raise ProtocolError(
                f"Nieznana akcja '{action_name}' dla aplikacji {self.APP_KEY}",
                {"available": sorted(self.actions())},
            )

        self._validate_params(handler, action_name, params)
        return self.connection.run(self._execute, handler, params)

    def _validate_params(
        self,
        handler: Callable[..., Any],
        action_name: str,
        params: dict[str, Any],
    ) -> None:
        signature = inspect.signature(handler)
        accepts_kwargs = any(
            parameter.kind is inspect.Parameter.VAR_KEYWORD
            for parameter in signature.parameters.values()
        )
        allowed = {
            name
            for name, parameter in signature.parameters.items()
            if name != "self"
            and parameter.kind
            in (
                inspect.Parameter.POSITIONAL_OR_KEYWORD,
                inspect.Parameter.KEYWORD_ONLY,
            )
        }

        if not accepts_kwargs:
            unexpected = set(params) - allowed
            if unexpected:
                raise ProtocolError(
                    f"Akcja '{action_name}' nie przyjmuje parametrow: "
                    f"{', '.join(sorted(unexpected))}",
                    {"allowed": sorted(allowed)},
                )

        required = {
            name
            for name, parameter in signature.parameters.items()
            if name != "self" and parameter.default is inspect.Parameter.empty
        }
        missing = required - set(params)
        if missing:
            raise ProtocolError(
                f"Akcja '{action_name}' wymaga parametrow: {', '.join(sorted(missing))}",
                {"required": sorted(required)},
            )

    def _execute(self, handler: Callable[..., Any], params: dict[str, Any]) -> Any:
        try:
            return to_python_result(handler(self, **params))
        except BridgeError:
            raise
        except com_error as exc:
            raise self.map_com_error(exc) from exc
        except FileNotFoundError as exc:
            raise DocumentNotFoundError(f"Nie znaleziono pliku: {exc}") from exc
        except (ValueError, IndexError, KeyError) as exc:
            raise InvalidReferenceError(str(exc) or type(exc).__name__) from exc
        except TypeError as exc:
            raise ProtocolError(f"Nieprawidlowe argumenty akcji: {exc}") from exc
        except AttributeError as exc:
            raise UnsupportedOperationError(
                f"Operacja niedostepna w tej wersji {self.DISPLAY_NAME}: {exc}"
            ) from exc

    def map_com_error(self, exc: BaseException) -> BridgeError:
        """Tlumaczy ``com_error`` na czytelny wyjatek Bridge."""
        args = getattr(exc, "args", ())
        hresult = args[0] if args else None
        description = _com_description(exc)
        suffix = f": {description}" if description else ""

        if hresult in _DISCONNECTED:
            self.connection.reset()
            return ComConnectionError(
                f"Utracono polaczenie z {self.DISPLAY_NAME} - aplikacja zostala "
                f"zamknieta lub przestala odpowiadac{suffix}"
            )
        if hresult in _BUSY:
            return ComConnectionError(
                f"{self.DISPLAY_NAME} jest zajety i odrzucil wywolanie - zamknij "
                f"otwarte okno dialogowe i sprobuj ponownie{suffix}"
            )
        if hresult in _BAD_INDEX:
            return InvalidReferenceError(
                f"Odwolanie poza zakresem w {self.DISPLAY_NAME}{suffix}"
            )
        if hresult in _NOT_SUPPORTED:
            return UnsupportedOperationError(
                f"Operacja nieobslugiwana przez zainstalowana wersje "
                f"{self.DISPLAY_NAME}{suffix}"
            )
        if hresult in _FILE_ERRORS:
            return DocumentNotFoundError(f"Plik niedostepny{suffix}")

        message = description or str(exc)
        return BridgeError(
            f"Blad {self.DISPLAY_NAME}: {message}",
            {"hresult": hex(hresult & 0xFFFFFFFF) if isinstance(hresult, int) else None},
        )

    @contextlib.contextmanager
    def alerts_suppressed(self) -> Iterator[None]:
        """Wylacza okna dialogowe Office (np. pytanie o nadpisanie pliku)."""
        app = self.app
        previous: Any = None
        try:
            previous = app.DisplayAlerts
            app.DisplayAlerts = self.ALERTS_OFF
        except Exception:  # noqa: BLE001 - PowerPoint bywa kapryzny przy starcie
            previous = None

        try:
            yield
        finally:
            if previous is not None:
                try:
                    app.DisplayAlerts = previous
                except Exception:  # noqa: BLE001
                    pass

    def resolve_existing_path(self, path: str) -> str:
        """Sciezka istniejacego pliku albo :class:`DocumentNotFoundError`."""
        try:
            return normalize_path(path, must_exist=True)
        except FileNotFoundError as exc:
            raise DocumentNotFoundError(f"Nie znaleziono pliku: {exc}") from exc

    def resolve_target_path(self, path: str) -> str:
        """Sciezka do zapisu - katalog nadrzedny musi istniec."""
        resolved = normalize_path(path)
        parent = os.path.dirname(resolved)
        if parent and not os.path.isdir(parent):
            raise DocumentNotFoundError(f"Katalog docelowy nie istnieje: {parent}")
        return resolved

    @staticmethod
    def require_index(value: Any, maximum: int, label: str) -> int:
        """Waliduje indeks 1-based uzywany przez kolekcje COM."""
        try:
            index = int(value)
        except (TypeError, ValueError) as exc:
            raise InvalidReferenceError(f"{label} musi byc liczba calkowita") from exc
        if index < 1 or index > maximum:
            raise InvalidReferenceError(
                f"{label} = {index} poza zakresem 1..{maximum}"
                if maximum
                else f"{label} = {index}, ale dokument jest pusty"
            )
        return index

    @action()
    def ping(self) -> dict[str, Any]:
        """Sprawdza, czy aplikacja odpowiada (wymusza polaczenie COM)."""
        app = self.app
        return {
            "app": self.APP_KEY,
            "name": to_python(getattr(app, "Name", self.DISPLAY_NAME)),
            "version": to_python(getattr(app, "Version", None)),
            "connected": True,
        }

    @action()
    def status(self) -> dict[str, Any]:
        """Zwraca stan polaczenia bez wymuszania startu aplikacji."""
        return self.connection.status()


def is_connection_error(exc: BaseException) -> bool:
    """Czy blad COM dotyczy polaczenia (a nie argumentow wywolania)."""
    args = getattr(exc, "args", ())
    hresult = args[0] if args else None
    return hresult in _DISCONNECTED or hresult in _BUSY


def to_python_result(value: Any) -> Any:
    """Rekurencyjnie czysci wynik akcji do postaci serializowalnej w JSON."""
    if isinstance(value, dict):
        return {str(key): to_python_result(item) for key, item in value.items()}
    if isinstance(value, list):
        return [to_python_result(item) for item in value]
    if isinstance(value, tuple):
        return [to_python_result(item) for item in value]
    return to_python(value)


def _com_description(exc: BaseException) -> str:
    """Wyciaga czytelny opis bledu z ``excepinfo`` wyjatku COM."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], (tuple, list)) and len(args[2]) >= 3:
        description = args[2][2]
        if description:
            return str(description).strip()
    if len(args) >= 2 and args[1]:
        return str(args[1]).strip()
    return ""
