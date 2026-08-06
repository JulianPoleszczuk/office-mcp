"""Hierarchia wyjatkow Bridge.

Kontrolery nigdy nie przepuszczaja surowego ``pywintypes.com_error`` do
warstwy MCP - kazdy blad COM jest mapowany na jeden z ponizszych typow,
ktore serializuja sie do stabilnego JSON-a (``type`` + ``message``).
"""

from __future__ import annotations


class BridgeError(Exception):
    """Bazowy wyjatek Bridge - wszystko co leci do klienta dziedziczy po nim."""

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
    """Aplikacja Office niedostepna, zamknieta w trakcie pracy albo nie odpowiada."""


class DocumentNotFoundError(BridgeError):
    """Plik nie istnieje albo zadany dokument nie jest otwarty."""


class InvalidReferenceError(BridgeError):
    """Zly indeks slajdu/akapitu, nieistniejacy arkusz, adres komorki poza zakresem."""


class UnsupportedOperationError(BridgeError):
    """Operacja nieobslugiwana przez dana wersje Office (np. brak funkcji w 2019)."""


class ProtocolError(BridgeError):
    """Niepoprawna wiadomosc na wejsciu Bridge (zly JSON, brak wymaganych pol)."""


class TimeoutError_(BridgeError):
    """Wywolanie COM przekroczylo limit czasu - aplikacja prawdopodobnie wisi."""

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
