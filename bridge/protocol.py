"""Protokol Bridge: JSON-line po TCP.

Jedna wiadomosc = jedna linia JSON zakonczona ``\\n``.

Request  (MCP Server -> Bridge)::

    {"id": "uuid", "app": "powerpoint", "action": "add_slide",
     "params": {"layout": "title_content", "title": "Wstep"}}

Response (Bridge -> MCP Server)::

    {"id": "uuid", "ok": true, "result": {"slide_index": 2}}
    {"id": "uuid", "ok": false,
     "error": {"type": "ComConnectionError", "message": "PowerPoint nie odpowiada"}}
"""

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass, field
from typing import Any

from bridge.utils.errors import BridgeError, ProtocolError

KNOWN_APPS = ("powerpoint", "excel", "word")

ENCODING = "utf-8"


@dataclass(slots=True)
class Request:
    """Pojedyncze zadanie wykonania akcji przez kontroler danej aplikacji."""

    app: str
    action: str
    params: dict[str, Any] = field(default_factory=dict)
    id: str = field(default_factory=lambda: str(uuid.uuid4()))

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "app": self.app,
            "action": self.action,
            "params": self.params,
        }

    def encode(self) -> bytes:
        return encode_line(self.to_dict())

    @classmethod
    def from_dict(cls, data: Any) -> "Request":
        if not isinstance(data, dict):
            raise ProtocolError("Request musi byc obiektem JSON")

        app = data.get("app")
        action = data.get("action")
        if not isinstance(app, str) or not app:
            raise ProtocolError("Brak wymaganego pola 'app'")
        if not isinstance(action, str) or not action:
            raise ProtocolError("Brak wymaganego pola 'action'")

        app = app.strip().lower()
        if app not in KNOWN_APPS:
            raise ProtocolError(
                f"Nieznana aplikacja '{app}', dozwolone: {', '.join(KNOWN_APPS)}"
            )

        params = data.get("params", {})
        if params is None:
            params = {}
        if not isinstance(params, dict):
            raise ProtocolError("Pole 'params' musi byc obiektem JSON")

        request_id = data.get("id")
        if request_id is not None and not isinstance(request_id, str):
            raise ProtocolError("Pole 'id' musi byc tekstem")

        return cls(
            app=app,
            action=action.strip(),
            params=params,
            id=request_id or str(uuid.uuid4()),
        )

    @classmethod
    def decode(cls, line: bytes | str) -> "Request":
        return cls.from_dict(decode_line(line))


@dataclass(slots=True)
class Response:
    """Odpowiedz Bridge - zawsze albo ``result``, albo ``error``."""

    id: str
    ok: bool
    result: Any = None
    error: dict[str, Any] | None = None

    @classmethod
    def success(cls, request_id: str, result: Any = None) -> "Response":
        return cls(id=request_id, ok=True, result=result)

    @classmethod
    def failure(cls, request_id: str, error: BridgeError | Exception) -> "Response":
        if isinstance(error, BridgeError):
            payload = error.to_dict()
        else:
            payload = {"type": type(error).__name__, "message": str(error)}
        return cls(id=request_id, ok=False, error=payload)

    def to_dict(self) -> dict[str, Any]:
        if self.ok:
            return {"id": self.id, "ok": True, "result": self.result}
        return {"id": self.id, "ok": False, "error": self.error}

    def encode(self) -> bytes:
        return encode_line(self.to_dict())

    @classmethod
    def from_dict(cls, data: Any) -> "Response":
        if not isinstance(data, dict):
            raise ProtocolError("Response musi byc obiektem JSON")
        if "ok" not in data:
            raise ProtocolError("Brak wymaganego pola 'ok'")

        ok = bool(data["ok"])
        error = data.get("error")
        if not ok and not isinstance(error, dict):
            raise ProtocolError("Odpowiedz bledu musi zawierac obiekt 'error'")

        return cls(
            id=str(data.get("id", "")),
            ok=ok,
            result=data.get("result"),
            error=error,
        )

    @classmethod
    def decode(cls, line: bytes | str) -> "Response":
        return cls.from_dict(decode_line(line))


def encode_line(payload: dict[str, Any]) -> bytes:
    """Serializuje slownik do jednej linii JSON zakonczonej znakiem nowej linii."""
    return (json.dumps(payload, ensure_ascii=False, default=str) + "\n").encode(ENCODING)


def decode_line(line: bytes | str) -> Any:
    """Parsuje jedna linie protokolu, mapujac bledy JSON na ProtocolError."""
    if isinstance(line, bytes):
        try:
            line = line.decode(ENCODING)
        except UnicodeDecodeError as exc:
            raise ProtocolError(f"Linia nie jest poprawnym UTF-8: {exc}") from exc

    line = line.strip()
    if not line:
        raise ProtocolError("Pusta linia protokolu")

    try:
        return json.loads(line)
    except json.JSONDecodeError as exc:
        raise ProtocolError(f"Niepoprawny JSON: {exc.msg} (pozycja {exc.pos})") from exc
