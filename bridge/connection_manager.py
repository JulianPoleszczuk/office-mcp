"""Managing COM connections to Office apps.

Rules this module enforces:

* **lazy connect** - COM starts on the first action for a given app, not when
  the Bridge starts,
* **attach to a running instance** - ``GetActiveObject`` first, ``Dispatch``
  only as a fallback, so the user does not get a second window,
* **failure isolation** - each app gets its own thread, its own state and its
  own connection object, so a hung Word does not block Excel,
* **timeout** - every COM call is time limited; once it is exceeded the thread
  is abandoned and the connection marked dead.

All COM calls for one app go through a single dedicated thread, because COM
objects from an STA apartment cannot be used from other threads.
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
    """Static description of an Office app the Bridge can drive."""

    key: str
    prog_id: str
    display_name: str


APP_SPECS: dict[str, AppSpec] = {
    "powerpoint": AppSpec("powerpoint", "PowerPoint.Application", "PowerPoint"),
    "excel": AppSpec("excel", "Excel.Application", "Excel"),
    "word": AppSpec("word", "Word.Application", "Word"),
}


def ensure_windows() -> None:
    """Hard guard - Office COM exists on Windows only."""
    if sys.platform != "win32":
        raise ComConnectionError(
            "office-mcp runs on Windows only - Office automation is built on "
            f"COM, which does not exist on platform '{sys.platform}'."
        )
    if not COM_AVAILABLE:
        raise ComConnectionError(
            "Brak pywin32. Zainstaluj zaleznosci: pip install -r requirements.txt "
            "then run python Scripts/pywin32_postinstall.py -install"
        )


class AppConnection:
    """A live COM connection to one Office app, with its own STA thread."""

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
        """Runs ``func`` on this app's COM thread, enforcing the time limit."""
        limit = timeout if timeout is not None else self.timeout
        executor = self._ensure_executor()
        future: Future = executor.submit(self._invoke, func, args)

        try:
            return future.result(timeout=limit)
        except FutureTimeout:
            self._abandon_executor()
            self._last_error = f"Przekroczono limit {limit:.0f}s"
            raise ComTimeoutError(
                f"{self.spec.display_name} did not answer within {limit:.0f}s - "
                "the app may be waiting on the user (an open dialog box?)."
            ) from None

    def _invoke(self, func: Callable[..., T], args: tuple[Any, ...]) -> T:
        try:
            return func(*args)
        except com_error as exc:
            if _is_disconnected(exc):
                self._drop_reference()
            raise

    def application(self) -> Any:
        """Returns the COM Application object; connects lazily on first use.

        Called only from this app's COM thread.
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
                "%s is not running - starting a new instance",
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
                    f"Could not start {self.spec.display_name}. Check that "
                    "Office is installed and that the app is not stuck in the "
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
            logger.debug("Could not set Visible on %s: %s", self.spec.key, exc)

    def _is_alive(self, app: Any) -> bool:
        try:
            _ = app.Name
            return True
        except Exception:  # noqa: BLE001
            return False

    def _drop_reference(self) -> None:
        self._app = None

    def reset(self) -> None:
        """Forgets the COM object - the next call reconnects from scratch."""
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
        """Abandons a stuck COM thread and assumes the connection is dead."""
        with self._lock:
            executor, self._executor = self._executor, None
            self._app = None
        if executor is not None:
            executor.shutdown(wait=False)

    def close(self) -> None:
        """Shuts the COM thread down (without closing the Office app itself)."""
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
            raise ComConnectionError(f"Unsupported app: {app_key!r}")

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
    """Initialises the COM apartment on a pool worker thread."""
    try:
        import pythoncom

        pythoncom.CoInitialize()
    except Exception as exc:  # noqa: BLE001 - thread keeps going, error surfaces at Dispatch
        logger.debug("CoInitialize nieudane: %s", exc)


_DISCONNECTED_HRESULTS = {
    -2147417848,  # RPC_E_DISCONNECTED / object vanished along with the app
    -2147023174,  # RPC server unavailable
    -2147023170,  # RPC failed
    -2146959355,  # CO_E_SERVER_EXEC_FAILURE
    -2147221021,  # CO_E_OBJNOTCONNECTED
    -2147221164,  # REGDB_E_CLASSNOTREG
}


def _is_disconnected(exc: BaseException) -> bool:
    """Whether a COM error means the connection to the app was lost."""
    hresult = getattr(exc, "args", (None,))[0]
    return hresult in _DISCONNECTED_HRESULTS
