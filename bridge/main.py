"""Proces Bridge - serwer TCP (JSON-line) trzymajacy zywe polaczenia COM.

Uruchomienie recznie (przydatne przy debugowaniu)::

    python -m bridge.main --port 8765 --log-level DEBUG

Bridge celowo jest osobnym, dlugozyjacym procesem: serwer MCP moze byc
restartowany przez klienta (Claude Desktop / Claude Code) bez zrywania
polaczen COM i bez zamykania dokumentow otwartych przez uzytkownika.
"""

from __future__ import annotations

import argparse
import logging
import os
import socketserver
import sys
import threading
from typing import Any

from bridge.connection_manager import APP_SPECS, ConnectionManager, ensure_windows
from bridge.controllers.excel import ExcelController
from bridge.controllers.powerpoint import PowerPointController
from bridge.controllers.word import WordController
from bridge.protocol import Request, Response
from bridge.utils.errors import BridgeError, ProtocolError

logger = logging.getLogger("bridge")

DEFAULT_HOST = os.environ.get("OFFICE_BRIDGE_HOST", "127.0.0.1")
DEFAULT_PORT = int(os.environ.get("OFFICE_BRIDGE_PORT", "8765"))
DEFAULT_TIMEOUT = float(os.environ.get("OFFICE_BRIDGE_TIMEOUT", "15"))

CONTROLLER_TYPES = {
    "powerpoint": PowerPointController,
    "excel": ExcelController,
    "word": WordController,
}


class Dispatcher:
    """Kieruje zadania do kontrolera wlasciwej aplikacji."""

    def __init__(self, timeout: float = DEFAULT_TIMEOUT) -> None:
        self.manager = ConnectionManager(timeout=timeout)
        self._controllers: dict[str, Any] = {}
        self._lock = threading.Lock()

    def controller(self, app_key: str) -> Any:
        with self._lock:
            controller = self._controllers.get(app_key)
            if controller is None:
                controller_type = CONTROLLER_TYPES.get(app_key)
                if controller_type is None:
                    raise ProtocolError(f"Nieobslugiwana aplikacja: {app_key!r}")
                controller = controller_type(self.manager.connection(app_key))
                self._controllers[app_key] = controller
            return controller

    def handle(self, request: Request) -> Response:
        """Wykonuje zadanie i zawsze zwraca poprawna odpowiedz protokolu."""
        try:
            result = self.controller(request.app).dispatch(request.action, request.params)
            return Response.success(request.id, result)
        except BridgeError as exc:
            logger.warning(
                "%s.%s -> %s: %s",
                request.app,
                request.action,
                exc.type_name,
                exc.message,
            )
            return Response.failure(request.id, exc)
        except Exception as exc:  # noqa: BLE001 - Bridge nigdy nie moze wywalic sie na kliencie
            logger.exception("Nieoczekiwany blad %s.%s", request.app, request.action)
            return Response.failure(request.id, exc)

    def shutdown(self) -> None:
        self.manager.shutdown()


class BridgeRequestHandler(socketserver.StreamRequestHandler):
    """Obsluga jednego polaczenia TCP - kolejne linie JSON az do rozlaczenia."""

    server: "BridgeServer"

    def handle(self) -> None:
        peer = self.client_address
        logger.info("Klient polaczony: %s", peer)
        try:
            for line in self.rfile:
                if not line.strip():
                    continue
                response = self._handle_line(line)
                self.wfile.write(response.encode())
                self.wfile.flush()
        except (ConnectionResetError, BrokenPipeError):
            logger.info("Klient rozlaczyl sie nagle: %s", peer)
        finally:
            logger.info("Klient rozlaczony: %s", peer)

    def _handle_line(self, line: bytes) -> Response:
        try:
            request = Request.decode(line)
        except ProtocolError as exc:
            logger.warning("Odrzucona wiadomosc: %s", exc.message)
            return Response.failure("", exc)

        logger.debug("-> %s.%s %s", request.app, request.action, request.params)
        response = self.server.dispatcher.handle(request)
        logger.debug("<- ok=%s", response.ok)
        return response


class BridgeServer(socketserver.ThreadingTCPServer):
    """Wielowatkowy serwer TCP; watki COM zyja w :class:`ConnectionManager`."""

    allow_reuse_address = True
    daemon_threads = True

    def __init__(
        self,
        host: str = DEFAULT_HOST,
        port: int = DEFAULT_PORT,
        timeout: float = DEFAULT_TIMEOUT,
    ) -> None:
        self.dispatcher = Dispatcher(timeout=timeout)
        super().__init__((host, port), BridgeRequestHandler)

    def server_close(self) -> None:
        super().server_close()
        self.dispatcher.shutdown()


def serve(
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    timeout: float = DEFAULT_TIMEOUT,
) -> None:
    """Startuje Bridge i blokuje watek do przerwania (Ctrl+C)."""
    server = BridgeServer(host=host, port=port, timeout=timeout)
    logger.info(
        "Bridge nasluchuje na %s:%s (aplikacje: %s, timeout COM: %.0fs)",
        host,
        port,
        ", ".join(APP_SPECS),
        timeout,
    )
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("Zatrzymywanie Bridge...")
    finally:
        server.shutdown()
        server.server_close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m bridge.main",
        description="Office Bridge - most COM miedzy serwerem MCP a Office 2019",
    )
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument(
        "--timeout",
        type=float,
        default=DEFAULT_TIMEOUT,
        help="Limit czasu pojedynczego wywolania COM w sekundach",
    )
    parser.add_argument(
        "--log-level",
        default=os.environ.get("OFFICE_BRIDGE_LOG_LEVEL", "INFO"),
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        stream=sys.stderr,
    )

    try:
        ensure_windows()
    except BridgeError as exc:
        print(exc.message, file=sys.stderr)
        return 1

    serve(host=args.host, port=args.port, timeout=args.timeout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
