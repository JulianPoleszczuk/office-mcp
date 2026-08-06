import re
from pathlib import Path

import pytest

import server
from bridge.controllers.excel import ExcelController
from bridge.controllers.powerpoint import PowerPointController
from bridge.controllers.word import WordController
from bridge.protocol import Response

CONTROLLERS = {
    "powerpoint": PowerPointController,
    "excel": ExcelController,
    "word": WordController,
}

CALL_PATTERN = re.compile(r'call_bridge\(\s*"(\w+)",\s*"(\w+)"')


def bridge_calls() -> list[tuple[str, str]]:
    source = Path(server.__file__).read_text(encoding="utf-8")
    return CALL_PATTERN.findall(source)


class FakeClient:
    def __init__(self, response=None, error=None):
        self.response = response
        self.error = error
        self.calls = []

    def call(self, app, action, params):
        self.calls.append((app, action, params))
        if self.error is not None:
            raise self.error
        return self.response

    def status(self):
        return {"connected": True}

    def close(self):
        pass


@pytest.fixture
def fake_client(monkeypatch):
    def install(response=None, error=None):
        client = FakeClient(response=response, error=error)
        monkeypatch.setattr(server, "client", client)
        return client

    return install


class TestToolRegistry:
    def test_every_tool_action_exists_in_controller(self):
        for app, action in bridge_calls():
            assert action in CONTROLLERS[app].actions(), f"{app}.{action}"

    def test_all_three_apps_are_covered(self):
        apps = {app for app, _ in bridge_calls()}
        assert apps == {"powerpoint", "excel", "word"}


class TestCallBridge:
    def test_success_is_wrapped(self, fake_client):
        fake_client(response=Response.success("1", {"slide_index": 2}))

        result = server.call_bridge("powerpoint", "add_slide", {"layout": "blank"})

        assert result == {"ok": True, "result": {"slide_index": 2}}

    def test_bridge_error_is_passed_through(self, fake_client):
        error = {"type": "ComConnectionError", "message": "PowerPoint nie odpowiada"}
        fake_client(response=Response(id="1", ok=False, error=error))

        result = server.call_bridge("powerpoint", "save", {})

        assert result == {"ok": False, "error": error}

    def test_unavailable_bridge_becomes_structured_error(self, fake_client):
        fake_client(error=server.BridgeUnavailable("brak procesu"))

        result = server.call_bridge("excel", "save", {})

        assert result["ok"] is False
        assert result["error"]["type"] == "BridgeUnavailable"

    def test_unexpected_exception_is_caught(self, fake_client):
        fake_client(error=RuntimeError("cos poszlo nie tak"))

        result = server.call_bridge("word", "save", {})

        assert result["ok"] is False
        assert result["error"]["type"] == "RuntimeError"

    def test_none_parameters_are_dropped(self, fake_client):
        client = fake_client(response=Response.success("1", None))

        server.call_bridge("powerpoint", "add_slide", {"layout": "blank", "index": None})

        assert client.calls[0][2] == {"layout": "blank"}

    def test_explicit_none_can_be_kept(self, fake_client):
        client = fake_client(response=Response.success("1", None))

        server.call_bridge(
            "excel",
            "set_cell",
            {"sheet": "Arkusz1", "cell_ref": "A1", "value": None},
            keep_none=("value",),
        )

        assert client.calls[0][2]["value"] is None


class TestTools:
    def test_tool_returns_structured_payload(self, fake_client):
        fake_client(response=Response.success("1", {"slide_count": 5}))

        assert server.ppt_get_presentation_info() == {
            "ok": True,
            "result": {"slide_count": 5},
        }

    def test_tool_forwards_arguments(self, fake_client):
        client = fake_client(response=Response.success("1", {}))

        server.xl_set_cell(sheet="Budzet", cell_ref="B2", value=1500)

        assert client.calls[0] == (
            "excel",
            "set_cell",
            {"sheet": "Budzet", "cell_ref": "B2", "value": 1500},
        )

    def test_office_status_queries_every_app(self, fake_client):
        client = fake_client(response=Response.success("1", {"connected": True}))

        result = server.office_status()

        assert set(result["result"]["apps"]) == {"powerpoint", "excel", "word"}
        assert [call[1] for call in client.calls] == ["status", "status", "status"]
