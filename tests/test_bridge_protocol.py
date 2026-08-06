import json
import socket
import threading

import pytest

from bridge.protocol import Request, Response, decode_line, encode_line
from bridge.utils.errors import ComConnectionError, ProtocolError


class TestRequest:
    def test_roundtrip(self):
        req = Request(app="powerpoint", action="add_slide", params={"layout": "blank"})
        decoded = Request.decode(req.encode())

        assert decoded.id == req.id
        assert decoded.app == "powerpoint"
        assert decoded.action == "add_slide"
        assert decoded.params == {"layout": "blank"}

    def test_generates_id_when_missing(self):
        req = Request.from_dict({"app": "excel", "action": "save"})
        assert req.id

    def test_normalizes_app_name(self):
        req = Request.from_dict({"app": "  PowerPoint ", "action": "save"})
        assert req.app == "powerpoint"

    def test_defaults_empty_params(self):
        assert Request.from_dict({"app": "word", "action": "save"}).params == {}
        assert (
            Request.from_dict({"app": "word", "action": "save", "params": None}).params
            == {}
        )

    @pytest.mark.parametrize(
        "payload",
        [
            {"action": "save"},
            {"app": "excel"},
            {"app": "", "action": "save"},
            {"app": "notepad", "action": "save"},
            {"app": "excel", "action": "save", "params": []},
            {"app": "excel", "action": "save", "id": 7},
            ["excel", "save"],
        ],
    )
    def test_rejects_invalid_payloads(self, payload):
        with pytest.raises(ProtocolError):
            Request.from_dict(payload)

    def test_rejects_malformed_json(self):
        with pytest.raises(ProtocolError):
            Request.decode(b'{"app": "excel", ')

    def test_rejects_empty_line(self):
        with pytest.raises(ProtocolError):
            Request.decode("   \n")


class TestResponse:
    def test_success_roundtrip(self):
        resp = Response.success("abc", {"slide_index": 2})
        decoded = Response.decode(resp.encode())

        assert decoded.ok is True
        assert decoded.result == {"slide_index": 2}
        assert decoded.error is None

    def test_failure_from_bridge_error(self):
        resp = Response.failure("abc", ComConnectionError("PowerPoint nie odpowiada"))
        payload = json.loads(resp.encode().decode("utf-8"))

        assert payload["ok"] is False
        assert payload["error"]["type"] == "ComConnectionError"
        assert payload["error"]["message"] == "PowerPoint nie odpowiada"

    def test_failure_from_plain_exception(self):
        resp = Response.failure("abc", ValueError("zly argument"))
        assert resp.error == {"type": "ValueError", "message": "zly argument"}

    def test_failure_carries_details(self):
        err = ComConnectionError("brak apki", {"app": "word"})
        assert Response.failure("x", err).error["details"] == {"app": "word"}

    def test_failure_requires_error_object(self):
        with pytest.raises(ProtocolError):
            Response.from_dict({"id": "x", "ok": False})


class TestLineCodec:
    def test_encode_ends_with_newline(self):
        assert encode_line({"a": 1}).endswith(b"\n")

    def test_encode_keeps_unicode(self):
        assert "Wstęp" in encode_line({"title": "Wstęp"}).decode("utf-8")

    def test_encode_falls_back_to_str_for_unknown_types(self):
        payload = json.loads(encode_line({"v": {1, 2}}).decode("utf-8"))
        assert isinstance(payload["v"], str)

    def test_decode_accepts_bytes_and_str(self):
        assert decode_line(b'{"a": 1}') == {"a": 1}
        assert decode_line('{"a": 1}\n') == {"a": 1}

    def test_decode_rejects_broken_utf8(self):
        with pytest.raises(ProtocolError):
            decode_line(b'{"a": "\xff\xfe"}')


class _EchoServer:
    """Minimalny serwer JSON-line uzywany do testu integracyjnego transportu."""

    def __init__(self, handler):
        self.handler = handler
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(1)
        self.port = self.sock.getsockname()[1]
        self.thread = threading.Thread(target=self._serve, daemon=True)
        self.thread.start()

    def _serve(self):
        conn, _ = self.sock.accept()
        with conn, conn.makefile("rwb") as stream:
            for line in stream:
                try:
                    request = Request.decode(line)
                except ProtocolError as exc:
                    stream.write(Response.failure("", exc).encode())
                else:
                    try:
                        result = self.handler(request)
                        stream.write(Response.success(request.id, result).encode())
                    except Exception as exc:  # noqa: BLE001
                        stream.write(Response.failure(request.id, exc).encode())
                stream.flush()

    def close(self):
        self.sock.close()


@pytest.fixture
def echo_server():
    servers = []

    def make(handler):
        server = _EchoServer(handler)
        servers.append(server)
        return server

    yield make
    for server in servers:
        server.close()


class TestTransportIntegration:
    def test_request_response_over_socket(self, echo_server):
        server = echo_server(lambda req: {"echo": req.params, "action": req.action})

        with socket.create_connection(("127.0.0.1", server.port), timeout=5) as sock:
            stream = sock.makefile("rwb")
            request = Request(app="excel", action="set_cell", params={"value": 42})
            stream.write(request.encode())
            stream.flush()
            response = Response.decode(stream.readline())

        assert response.ok is True
        assert response.id == request.id
        assert response.result == {"echo": {"value": 42}, "action": "set_cell"}

    def test_multiple_requests_on_one_connection(self, echo_server):
        server = echo_server(lambda req: req.params.get("n"))
        ids = []

        with socket.create_connection(("127.0.0.1", server.port), timeout=5) as sock:
            stream = sock.makefile("rwb")
            for n in range(5):
                request = Request(app="word", action="ping", params={"n": n})
                ids.append(request.id)
                stream.write(request.encode())
                stream.flush()
                response = Response.decode(stream.readline())
                assert response.ok is True
                assert response.id == ids[n]
                assert response.result == n

    def test_error_travels_as_structured_payload(self, echo_server):
        def handler(_request):
            raise ComConnectionError("Word nie odpowiada")

        server = echo_server(handler)

        with socket.create_connection(("127.0.0.1", server.port), timeout=5) as sock:
            stream = sock.makefile("rwb")
            stream.write(Request(app="word", action="save").encode())
            stream.flush()
            response = Response.decode(stream.readline())

        assert response.ok is False
        assert response.error["type"] == "ComConnectionError"

    def test_malformed_line_does_not_kill_connection_echo(self, echo_server):
        server = echo_server(lambda req: "ok")

        with socket.create_connection(("127.0.0.1", server.port), timeout=5) as sock:
            stream = sock.makefile("rwb")
            stream.write(b"to nie jest json\n")
            stream.flush()
            broken = Response.decode(stream.readline())

            stream.write(Request(app="excel", action="ping").encode())
            stream.flush()
            good = Response.decode(stream.readline())

        assert broken.ok is False
        assert broken.error["type"] == "ProtocolError"
        assert good.ok is True


class _FakeController:
    """Kontroler bez COM - sprawdza sama sciezke zadanie -> odpowiedz w Bridge."""

    def __init__(self, connection):
        self.connection = connection

    def dispatch(self, action_name, params):
        if action_name == "com_down":
            raise ComConnectionError("Excel nie odpowiada")
        if action_name == "boom":
            raise RuntimeError("nieoczekiwany blad kontrolera")
        if action_name == "nieznana":
            raise ProtocolError("Nieznana akcja 'nieznana'")
        return {"action": action_name, "params": params}


@pytest.fixture
def bridge_server(monkeypatch):
    from bridge import main as bridge_main

    monkeypatch.setattr(
        bridge_main,
        "CONTROLLER_TYPES",
        {app: _FakeController for app in ("powerpoint", "excel", "word")},
    )

    server = bridge_main.BridgeServer(host="127.0.0.1", port=0)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    yield server

    server.shutdown()
    server.server_close()


def send(server, request: Request) -> Response:
    with socket.create_connection(server.server_address, timeout=5) as sock:
        stream = sock.makefile("rwb")
        stream.write(request.encode())
        stream.flush()
        return Response.decode(stream.readline())


class TestBridgeServer:
    def test_routes_request_to_controller(self, bridge_server):
        response = send(
            bridge_server, Request(app="excel", action="set_cell", params={"value": 7})
        )

        assert response.ok is True
        assert response.result == {"action": "set_cell", "params": {"value": 7}}

    def test_bridge_error_becomes_structured_response(self, bridge_server):
        response = send(bridge_server, Request(app="excel", action="com_down"))

        assert response.ok is False
        assert response.error["type"] == "ComConnectionError"
        assert "nie odpowiada" in response.error["message"]

    def test_unexpected_exception_does_not_crash_bridge(self, bridge_server):
        failed = send(bridge_server, Request(app="word", action="boom"))
        recovered = send(bridge_server, Request(app="word", action="save"))

        assert failed.ok is False
        assert failed.error["type"] == "RuntimeError"
        assert recovered.ok is True

    def test_unknown_action_is_reported_as_protocol_error(self, bridge_server):
        response = send(bridge_server, Request(app="powerpoint", action="nieznana"))

        assert response.ok is False
        assert response.error["type"] == "ProtocolError"

    def test_unknown_app_is_rejected_before_dispatch(self, bridge_server):
        with socket.create_connection(bridge_server.server_address, timeout=5) as sock:
            stream = sock.makefile("rwb")
            stream.write(b'{"id": "1", "app": "notepad", "action": "save"}\n')
            stream.flush()
            response = Response.decode(stream.readline())

        assert response.ok is False
        assert response.error["type"] == "ProtocolError"

    def test_controller_instances_are_reused_per_app(self, bridge_server):
        send(bridge_server, Request(app="word", action="save"))
        send(bridge_server, Request(app="word", action="save"))

        assert len(bridge_server.dispatcher._controllers) == 1
