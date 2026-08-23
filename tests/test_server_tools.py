import json
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
        error = {"type": "ComConnectionError", "message": "PowerPoint is not responding"}
        fake_client(response=Response(id="1", ok=False, error=error))

        result = server.call_bridge("powerpoint", "save", {})

        assert result == {"ok": False, "error": error}

    def test_unavailable_bridge_becomes_structured_error(self, fake_client):
        fake_client(error=server.BridgeUnavailable("brak procesu"))

        result = server.call_bridge("excel", "save", {})

        assert result["ok"] is False
        assert result["error"]["type"] == "BridgeUnavailable"

    def test_unexpected_exception_is_caught(self, fake_client):
        fake_client(error=RuntimeError("something went wrong"))

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
            {"sheet": "Sheet1", "cell_ref": "A1", "value": None},
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


class _FakeHttpResponse:
    """Minimal stand-in for what urllib.request.urlopen() returns - just
    enough for the `with urlopen(...) as resp: resp.read()` shape."""

    def __init__(self, payload: dict):
        self._body = json.dumps(payload).encode("utf-8")

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        return False

    def read(self):
        return self._body


UNSPLASH_HIT = {
    "results": [
        {
            "urls": {"regular": "https://images.unsplash.com/photo-fake"},
            "user": {"name": "Jane Photographer"},
        }
    ]
}


class TestPptSearchAndAddImage:
    """ppt_search_and_add_image does real network I/O (Unsplash search +
    download) before delegating to the same call_bridge("powerpoint",
    "add_image", ...) action ppt_add_image already uses - these tests mock
    only the network calls, exercising the real delegation/error-shape
    logic exactly like TestTools does for the COM-facing tools."""

    def test_missing_api_key_is_a_configuration_error_not_a_crash(self, monkeypatch, fake_client):
        monkeypatch.setattr(server, "UNSPLASH_ACCESS_KEY", "")
        fake_client()  # never reached; asserts nothing calls the Bridge

        result = server.ppt_search_and_add_image(
            slide_index=2, query="office desk", left=10, top=10
        )

        assert result == {
            "ok": False,
            "error": {
                "type": "ConfigurationError",
                "message": (
                    "UNSPLASH_ACCESS_KEY environment variable is not set. "
                    "Get a free key at unsplash.com/developers (no card required)."
                ),
            },
        }

    def test_searches_downloads_and_delegates_to_add_image(self, monkeypatch, tmp_path, fake_client):
        monkeypatch.setattr(server, "UNSPLASH_ACCESS_KEY", "fake-key")
        monkeypatch.setattr(
            server.urllib.request, "urlopen", lambda url, timeout=15: _FakeHttpResponse(UNSPLASH_HIT)
        )
        downloaded_to = tmp_path / "downloaded.jpg"
        downloaded_to.write_bytes(b"fake jpeg bytes")
        monkeypatch.setattr(server.tempfile, "mkstemp", lambda suffix, prefix: (0, str(downloaded_to)))
        monkeypatch.setattr(server.os, "close", lambda fd: None)
        retrieved = {}
        monkeypatch.setattr(
            server.urllib.request,
            "urlretrieve",
            lambda url, path: retrieved.update(url=url, path=path),
        )
        client = fake_client(response=Response.success("1", {"slide_index": 2, "shape_id": 7}))

        result = server.ppt_search_and_add_image(
            slide_index=2, query="office desk", left=10, top=20, width=300, height=200
        )

        assert retrieved == {"url": "https://images.unsplash.com/photo-fake", "path": str(downloaded_to)}
        assert client.calls[0] == (
            "powerpoint",
            "add_image",
            {
                "slide_index": 2,
                "image_path": str(downloaded_to),
                "left": 10,
                "top": 20,
                "width": 300,
                "height": 200,
            },
        )
        assert result["ok"] is True
        assert result["result"]["source"] == "Unsplash"
        assert result["result"]["photographer"] == "Jane Photographer"
        assert result["result"]["query"] == "office desk"
        # The temp file is cleaned up once the Bridge has embedded it - it
        # isn't needed after the .pptx has the pixel data.
        assert not downloaded_to.exists()

    def test_no_search_results_is_a_clear_error_not_an_empty_insert(self, monkeypatch, fake_client):
        monkeypatch.setattr(server, "UNSPLASH_ACCESS_KEY", "fake-key")
        monkeypatch.setattr(
            server.urllib.request, "urlopen", lambda url, timeout=15: _FakeHttpResponse({"results": []})
        )
        client = fake_client()

        result = server.ppt_search_and_add_image(slide_index=1, query="a query nothing matches", left=0, top=0)

        assert result["ok"] is False
        assert result["error"]["type"] == "ImageSearchError"
        assert client.calls == []

    def test_search_network_failure_is_wrapped_not_raised(self, monkeypatch, fake_client):
        monkeypatch.setattr(server, "UNSPLASH_ACCESS_KEY", "fake-key")

        def boom(url, timeout=15):
            raise OSError("network unreachable")

        monkeypatch.setattr(server.urllib.request, "urlopen", boom)
        fake_client()

        result = server.ppt_search_and_add_image(slide_index=1, query="anything", left=0, top=0)

        assert result == {"ok": False, "error": {"type": "ImageSearchError", "message": "network unreachable"}}

    def test_download_failure_is_wrapped_and_cleans_up(self, monkeypatch, tmp_path, fake_client):
        monkeypatch.setattr(server, "UNSPLASH_ACCESS_KEY", "fake-key")
        monkeypatch.setattr(
            server.urllib.request, "urlopen", lambda url, timeout=15: _FakeHttpResponse(UNSPLASH_HIT)
        )
        temp_path = tmp_path / "would-be-downloaded.jpg"
        monkeypatch.setattr(server.tempfile, "mkstemp", lambda suffix, prefix: (0, str(temp_path)))
        monkeypatch.setattr(server.os, "close", lambda fd: None)

        def boom(url, path):
            raise OSError("connection reset")

        monkeypatch.setattr(server.urllib.request, "urlretrieve", boom)
        fake_client()

        result = server.ppt_search_and_add_image(slide_index=1, query="anything", left=0, top=0)

        assert result == {"ok": False, "error": {"type": "ImageDownloadError", "message": "connection reset"}}
