"""Hierarchia wyjatkow Bridge.

Controllers never let a raw ``pywintypes.com_error`` reach the MCP layer -
every COM error is mapped to one of the types below, which serialise into
stable JSON (``type`` + ``message``).
"""

from __future__ import annotations


class BridgeError(Exception):
    """Base Bridge exception - everything sent to the client inherits from it."""

    def __init__(self, message: str, details: dict | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.details = details or {}

    @property
    def type_name(self) -> str:
        return type(self).__name__

    def to_dict(self) -> dict:
        payload = {"type": self.type_name, "message": self.message}
        if self.details:
            payload["details"] = self.details
        return payload


class ComConnectionError(BridgeError):
    """Office app unavailable, closed mid-operation or not responding."""


class DocumentNotFoundError(BridgeError):
    """File does not exist, or the requested document is not open."""


class InvalidReferenceError(BridgeError):
    """Bad slide/paragraph index, missing sheet, cell address out of range."""


class UnsupportedOperationError(BridgeError):
    """Operation unsupported by this Office version (e.g. missing in 2019)."""


class ProtocolError(BridgeError):
    """Niepoprawna wiadomosc na wejsciu Bridge (zly JSON, brak wymaganych pol)."""


class TimeoutError_(BridgeError):
    """A COM call exceeded its time limit - the app is probably hung."""

    @property
    def type_name(self) -> str:
        return "ComTimeoutError"


ComTimeoutError = TimeoutError_

ERROR_TYPES: dict[str, type[BridgeError]] = {
    "BridgeError": BridgeError,
    "ComConnectionError": ComConnectionError,
    "DocumentNotFoundError": DocumentNotFoundError,
    "InvalidReferenceError": InvalidReferenceError,
    "UnsupportedOperationError": UnsupportedOperationError,
    "ProtocolError": ProtocolError,
    "ComTimeoutError": ComTimeoutError,
}
