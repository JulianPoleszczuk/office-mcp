"""Zarzadzanie polaczeniami COM do aplikacji Office.

Zasady, ktore wymusza ten modul:

* **leniwe laczenie** - COM startuje dopiero przy pierwszej akcji dla danej
  aplikacji, nie przy starcie Bridge,
* **podlaczanie sie do otwartej instancji** - najpierw ``GetActiveObject``,
  dopiero potem ``Dispatch`` (czyli uzytkownik nie dostaje drugiego okna),
* **izolacja awarii** - kazda aplikacja ma wlasny watek, wlasny stan i wlasny
  obiekt polaczenia, wiec zawieszony Word nie blokuje Excela,
* **timeout** - kazde wywolanie COM ma limit czasu; po jego przekroczeniu
  watek jest porzucany, a polaczenie oznaczane jako martwe.

Wszystkie wywolania COM dla jednej aplikacji ida przez jeden dedykowany watek,
bo obiekty COM z apartamentu STA nie moga byc uzywane z innych watkow.
"""

from __future__ import annotations

import logging
import sys
import threading
from concurrent.futures import Future, ThreadPoolExecutor, TimeoutError as FutureTimeout
from dataclasses import dataclass
from typing import Any, Callable, TypeVar

from bridge.utils.com_helpers import COM_AVAILABLE, com_error
from bridge.utils.errors import ComConnectionError, ComTimeoutError

logger = logging.getLogger(__name__)

T = TypeVar("T")

DEFAULT_TIMEOUT = 15.0


@dataclass(frozen=True, slots=True)
class AppSpec:
    """Statyczny opis aplikacji Office obslugiwanej przez Bridge."""

    key: str
    prog_id: str
    display_name: str


APP_SPECS: dict[str, AppSpec] = {
    "powerpoint": AppSpec("powerpoint", "PowerPoint.Application", "PowerPoint"),
    "excel": AppSpec("excel", "Excel.Application", "Excel"),
    "word": AppSpec("word", "Word.Application", "Word"),
}


def ensure_windows() -> None:
    """Twardy guard - COM Office istnieje wylacznie na Windows."""
    if sys.platform != "win32":
        raise ComConnectionError(
            "office-mcp dziala tylko na Windows - automatyzacja Office opiera sie "
            f"na COM, ktorego nie ma na platformie '{sys.platform}'."
        )
    if not COM_AVAILABLE:
        raise ComConnectionError(
            "Brak pywin32. Zainstaluj zaleznosci: pip install -r requirements.txt "
            "oraz uruchom python Scripts/pywin32_postinstall.py -install"
        )


class AppConnection:
    """Zywe polaczenie COM do jednej aplikacji Office wraz z jej watkiem STA."""

    def __init__(self, spec: AppSpec, timeout: float = DEFAULT_TIMEOUT) -> None:
        self.spec = spec
        self.timeout = timeout
        self._app: Any = None
        self._lock = threading.Lock()
        self._executor: ThreadPoolExecutor | None = None
        self._launched_by_bridge = False
        self._last_error: str | None = None

    @property
    def is_connected(self) -> bool:
        return self._app is not None

    @property
    def launched_by_bridge(self) -> bool:
        return self._launched_by_bridge

    def status(self) -> dict[str, Any]:
        return {
            "app": self.spec.key,
            "display_name": self.spec.display_name,
            "connected": self.is_connected,
            "launched_by_bridge": self._launched_by_bridge,
            "last_error": self._last_error,
        }

    def run(self, func: Callable[..., T], *args: Any, timeout: float | None = None) -> T:
        """Wykonuje ``func`` w watku COM tej aplikacji, pilnujac limitu czasu."""
        limit = timeout if timeout is not None else self.timeout
        executor = self._ensure_executor()
        future: Future = executor.submit(self._invoke, func, args)

        try:
            return future.result(timeout=limit)
        except FutureTimeout:
            self._abandon_executor()
            self._last_error = f"Przekroczono limit {limit:.0f}s"
            raise ComTimeoutError(
                f"{self.spec.display_name} nie odpowiedzial w ciagu {limit:.0f}s - "
                "aplikacja moze czekac na akcje uzytkownika (otwarte okno dialogowe?)."
            ) from None

    def _invoke(self, func: Callable[..., T], args: tuple[Any, ...]) -> T:
        try:
            return func(*args)
        except com_error as exc:
            if _is_disconnected(exc):
                self._drop_reference()
            raise

    def application(self) -> Any:
        """Zwraca obiekt aplikacji COM; laczy sie leniwie przy pierwszym uzyciu.

        Wolane wylacznie z watku COM danej aplikacji.
        """
        if self._app is not None:
            if self._is_alive(self._app):
                return self._app
            self._drop_reference()
        return self._connect()

    def _connect(self) -> Any:
        ensure_windows()
        import win32com.client

        app = None
        try:
            app = win32com.client.GetActiveObject(self.spec.prog_id)
            self._launched_by_bridge = False
            logger.info("Podlaczono do otwartej instancji %s", self.spec.display_name)
        except com_error:
            logger.info(
                "%s nie jest otwarty - uruchamiam nowa instancje",
                self.spec.display_name,
            )
        except Exception as exc:  # noqa: BLE001
            logger.debug("GetActiveObject(%s) nieudane: %s", self.spec.prog_id, exc)

        if app is None:
            try:
                app = win32com.client.Dispatch(self.spec.prog_id)
                self._launched_by_bridge = True
            except com_error as exc:
                self._last_error = str(exc)
                raise ComConnectionError(
                    f"Nie udalo sie uruchomic {self.spec.display_name}. "
                    "Sprawdz, czy Office jest zainstalowany i czy aplikacja nie "
                    "wisi w tle (Menedzer zadan).",
                    {"prog_id": self.spec.prog_id},
                ) from exc

        self._make_visible(app)
        self._app = app
        self._last_error = None
        return app

    def _make_visible(self, app: Any) -> None:
        try:
            app.Visible = True
        except Exception as exc:  # noqa: BLE001
            logger.debug("Nie udalo sie ustawic Visible dla %s: %s", self.spec.key, exc)

    def _is_alive(self, app: Any) -> bool:
        try:
            _ = app.Name
            return True
        except Exception:  # noqa: BLE001
            return False

    def _drop_reference(self) -> None:
        self._app = None

    def reset(self) -> None:
        """Zapomina obiekt COM - kolejne wywolanie polaczy sie od nowa."""
        self._drop_reference()

    def _ensure_executor(self) -> ThreadPoolExecutor:
        with self._lock:
            if self._executor is None:
                self._executor = ThreadPoolExecutor(
                    max_workers=1,
                    thread_name_prefix=f"com-{self.spec.key}",
                    initializer=_init_com_apartment,
                )
            return self._executor

    def _abandon_executor(self) -> None:
        """Porzuca zablokowany watek COM i zaklada, ze polaczenie jest martwe."""
        with self._lock:
            executor, self._executor = self._executor, None
            self._app = None
        if executor is not None:
            executor.shutdown(wait=False)

    def close(self) -> None:
        """Zamyka watek COM (bez zamykania samej aplikacji Office)."""
        with self._lock:
            executor, self._executor = self._executor, None
            self._app = None
        if executor is not None:
            executor.shutdown(wait=False)


class ConnectionManager:
    """Rejestr polaczen - po jednym :class:`AppConnection` na aplikacje Office."""

    def __init__(self, timeout: float = DEFAULT_TIMEOUT) -> None:
        self.timeout = timeout
        self._connections: dict[str, AppConnection] = {}
        self._lock = threading.Lock()

    def connection(self, app_key: str) -> AppConnection:
        key = app_key.strip().lower()
        spec = APP_SPECS.get(key)
        if spec is None:
            raise ComConnectionError(f"Nieobslugiwana aplikacja: {app_key!r}")

        with self._lock:
            connection = self._connections.get(key)
            if connection is None:
                connection = AppConnection(spec, timeout=self.timeout)
                self._connections[key] = connection
            return connection

    def status(self) -> dict[str, Any]:
        with self._lock:
            known = dict(self._connections)
        return {
            key: known[key].status()
            if key in known
            else {
                "app": key,
                "display_name": spec.display_name,
                "connected": False,
                "launched_by_bridge": False,
                "last_error": None,
            }
            for key, spec in APP_SPECS.items()
        }

    def reset(self, app_key: str) -> None:
        self.connection(app_key).reset()

    def shutdown(self) -> None:
        with self._lock:
            connections = list(self._connections.values())
            self._connections.clear()
        for connection in connections:
            connection.close()


def _init_com_apartment() -> None:
    """Inicjalizuje apartament COM w watku roboczym puli."""
    try:
        import pythoncom

        pythoncom.CoInitialize()
    except Exception as exc:  # noqa: BLE001 - watek dziala dalej, blad wyjdzie przy Dispatch
        logger.debug("CoInitialize nieudane: %s", exc)


_DISCONNECTED_HRESULTS = {
    -2147417848,  # RPC_E_DISCONNECTED / obiekt zniknal razem z aplikacja
    -2147023174,  # RPC server unavailable
    -2147023170,  # RPC failed
    -2146959355,  # CO_E_SERVER_EXEC_FAILURE
    -2147221021,  # CO_E_OBJNOTCONNECTED
    -2147221164,  # REGDB_E_CLASSNOTREG
}


def _is_disconnected(exc: BaseException) -> bool:
    """Sprawdza, czy blad COM oznacza utrate polaczenia z aplikacja."""
    hresult = getattr(exc, "args", (None,))[0]
    return hresult in _DISCONNECTED_HRESULTS
