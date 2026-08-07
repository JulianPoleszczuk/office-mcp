"""Wspolna logika kontrolerow COM.

A controller is a thin layer over one Office app's ``Application`` object.
Office app. Methods marked with the :func:`action` decorator become Bridge
protokolu Bridge - reszta klasy to zwykle helpery.

Every action runs on its app's COM thread (see
:mod:`bridge.connection_manager`), and every exception - including raw
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
    """Marks a controller method as an action reachable over the Bridge protocol."""

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
    -2147418111,  # RPC_E_CALL_REJECTED - app busy (e.g. a dialog is open)
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
    """Controller base - action routing, error mapping, shared helpers."""

    APP_KEY: str = ""
    DISPLAY_NAME: str = ""
    ALERTS_OFF: Any = False

    def __init__(self, connection: AppConnection) -> None:
        self.connection = connection

    @classmethod
    def actions(cls) -> dict[str, Callable[..., Any]]:
        """Returns an ``action name -> method`` map built from the decorators."""
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
        """The ``Application`` object - connects lazily on first use."""
        return self.connection.application()

    def dispatch(self, action_name: str, params: dict[str, Any]) -> Any:
        """Runs the action on the app's COM thread and returns JSON-ready output."""
        handler = self.actions().get(action_name)
        if handler is None:
            raise ProtocolError(
                f"Unknown action '{action_name}' for app {self.APP_KEY}",
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
                    f"Action '{action_name}' does not accept parameters: "
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
                f"Action '{action_name}' requires parameters: {', '.join(sorted(missing))}",
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
            raise DocumentNotFoundError(f"File not found: {exc}") from exc
        except (ValueError, IndexError, KeyError) as exc:
            raise InvalidReferenceError(str(exc) or type(exc).__name__) from exc
        except TypeError as exc:
            raise ProtocolError(f"Invalid action arguments: {exc}") from exc
        except AttributeError as exc:
            raise UnsupportedOperationError(
                f"Operation unavailable in this version of {self.DISPLAY_NAME}: {exc}"
            ) from exc

    def map_com_error(self, exc: BaseException) -> BridgeError:
        """Translates a ``com_error`` into a readable Bridge exception."""
        args = getattr(exc, "args", ())
        hresult = args[0] if args else None
        description = _com_description(exc)
        suffix = f": {description}" if description else ""

        if hresult in _DISCONNECTED:
            self.connection.reset()
            return ComConnectionError(
                f"Lost connection to {self.DISPLAY_NAME} - the app was closed "
                f"or stopped responding{suffix}"
            )
        if hresult in _BUSY:
            return ComConnectionError(
                f"{self.DISPLAY_NAME} is busy and rejected the call - close any "
                f"open dialog box and try again{suffix}"
            )
        if hresult in _BAD_INDEX:
            return InvalidReferenceError(
                f"Reference out of range in {self.DISPLAY_NAME}{suffix}"
            )
        if hresult in _NOT_SUPPORTED:
            return UnsupportedOperationError(
                f"Operation not supported by the installed version of "
                f"{self.DISPLAY_NAME}{suffix}"
            )
        if hresult in _FILE_ERRORS:
            return DocumentNotFoundError(f"File unavailable{suffix}")

        message = description or str(exc)
        return BridgeError(
            f"{self.DISPLAY_NAME} error: {message}",
            {"hresult": hex(hresult & 0xFFFFFFFF) if isinstance(hresult, int) else None},
        )

    @contextlib.contextmanager
    def alerts_suppressed(self) -> Iterator[None]:
        """Turns off Office dialog boxes (e.g. the overwrite prompt)."""
        app = self.app
        previous: Any = None
        try:
            previous = app.DisplayAlerts
            app.DisplayAlerts = self.ALERTS_OFF
        except Exception:  # noqa: BLE001 - PowerPoint can be temperamental at startup
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
        """Path to an existing file, or :class:`DocumentNotFoundError`."""
        try:
            return normalize_path(path, must_exist=True)
        except FileNotFoundError as exc:
            raise DocumentNotFoundError(f"File not found: {exc}") from exc

    def resolve_target_path(self, path: str) -> str:
        """Path to save to - the parent directory must already exist."""
        resolved = normalize_path(path)
        parent = os.path.dirname(resolved)
        if parent and not os.path.isdir(parent):
            raise DocumentNotFoundError(f"Target directory does not exist: {parent}")
        return resolved

    @staticmethod
    def require_index(value: Any, maximum: int, label: str) -> int:
        """Validates the 1-based index used by COM collections."""
        try:
            index = int(value)
        except (TypeError, ValueError) as exc:
            raise InvalidReferenceError(f"{label} must be a whole number") from exc
        if index < 1 or index > maximum:
            raise InvalidReferenceError(
                f"{label} = {index} is outside the range 1..{maximum}"
                if maximum
                else f"{label} = {index}, but the document is empty"
            )
        return index

    @action()
    def ping(self) -> dict[str, Any]:
        """Checks that the app responds (forces the COM connection)."""
        app = self.app
        return {
            "app": self.APP_KEY,
            "name": to_python(getattr(app, "Name", self.DISPLAY_NAME)),
            "version": to_python(getattr(app, "Version", None)),
            "connected": True,
        }

    @action()
    def status(self) -> dict[str, Any]:
        """Returns connection state without forcing the app to start."""
        return self.connection.status()


def is_connection_error(exc: BaseException) -> bool:
    """Whether a COM error is about the connection (not the call arguments)."""
    args = getattr(exc, "args", ())
    hresult = args[0] if args else None
    return hresult in _DISCONNECTED or hresult in _BUSY


def to_python_result(value: Any) -> Any:
    """Recursively cleans an action result into a JSON-serialisable form."""
    if isinstance(value, dict):
        return {str(key): to_python_result(item) for key, item in value.items()}
    if isinstance(value, list):
        return [to_python_result(item) for item in value]
    if isinstance(value, tuple):
        return [to_python_result(item) for item in value]
    return to_python(value)


def _com_description(exc: BaseException) -> str:
    """Pulls a readable error description out of a COM exception's ``excepinfo``."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], (tuple, list)) and len(args[2]) >= 3:
        description = args[2][2]
        if description:
            return str(description).strip()
    if len(args) >= 2 and args[1]:
        return str(args[1]).strip()
    return ""
