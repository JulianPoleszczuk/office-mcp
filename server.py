"""MCP server for Microsoft Office 2019 (PowerPoint, Excel, Word).

The MCP layer is thin: every tool turns its arguments into a Bridge protocol
request and returns structured JSON::

    {"ok": true,  "result": {...}}
    {"ok": false, "error": {"type": "ComConnectionError", "message": "..."}}

No tool ever propagates a Python exception - an Office crash, a missing Bridge
or a bad sheet name always come back as a described error.

Running it (stdio, the way Claude Desktop starts it)::

    python server.py

The Bridge (the process holding COM connections) starts automatically on the
first tool call. You can also start it by hand:

    python -m bridge.main --log-level DEBUG
"""

from __future__ import annotations

import logging
import os
import socket
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Any

from bridge.protocol import Request, Response

try:  # mcp >= 2.0
    from mcp.server.mcpserver import MCPServer as _McpServer
except ImportError:  # pragma: no cover - mcp 1.x
    from mcp.server.fastmcp import FastMCP as _McpServer  # type: ignore[no-redef]

logger = logging.getLogger("office-mcp")

ROOT = Path(__file__).resolve().parent

BRIDGE_HOST = os.environ.get("OFFICE_BRIDGE_HOST", "127.0.0.1")
BRIDGE_PORT = int(os.environ.get("OFFICE_BRIDGE_PORT", "8765"))
BRIDGE_AUTOSTART = os.environ.get("OFFICE_BRIDGE_AUTOSTART", "1") not in ("0", "false", "no")
BRIDGE_START_TIMEOUT = float(os.environ.get("OFFICE_BRIDGE_START_TIMEOUT", "25"))
CALL_TIMEOUT = float(os.environ.get("OFFICE_MCP_CALL_TIMEOUT", "60"))

CREATE_NO_WINDOW = 0x08000000
DETACHED_PROCESS = 0x00000008


class BridgeUnavailable(RuntimeError):
    """The Bridge is not answering - it could not be started or reached."""


class BridgeClient:
    """Bridge TCP client: one connection, process autostart, reconnect."""

    def __init__(
        self,
        host: str = BRIDGE_HOST,
        port: int = BRIDGE_PORT,
        autostart: bool = BRIDGE_AUTOSTART,
        call_timeout: float = CALL_TIMEOUT,
    ) -> None:
        self.host = host
        self.port = port
        self.autostart = autostart
        self.call_timeout = call_timeout
        self._lock = threading.Lock()
        self._socket: socket.socket | None = None
        self._stream: Any = None
        self._process: subprocess.Popen | None = None

    def call(self, app: str, action: str, params: dict[str, Any]) -> Response:
        """Sends a request to the Bridge; retries once on a dropped connection."""
        request = Request(app=app, action=action, params=params)
        last_error: Exception | None = None

        with self._lock:
            for attempt in (1, 2):
                try:
                    stream = self._connected_stream()
                    stream.write(request.encode())
                    stream.flush()
                    line = stream.readline()
                    if not line:
                        raise BridgeUnavailable("The Bridge closed the connection")
                    return Response.decode(line)
                except (OSError, BridgeUnavailable) as exc:
                    last_error = exc
                    self._disconnect()
                    if attempt == 2:
                        break

        raise BridgeUnavailable(
            f"No connection to the Bridge at {self.host}:{self.port} ({last_error})"
        )

    def status(self) -> dict[str, Any]:
        return {
            "host": self.host,
            "port": self.port,
            "connected": self._socket is not None,
            "autostart": self.autostart,
            "bridge_process_pid": self._process.pid if self._process else None,
        }

    def close(self) -> None:
        with self._lock:
            self._disconnect()

    def _connected_stream(self) -> Any:
        if self._stream is not None:
            return self._stream

        try:
            self._open_socket()
        except OSError:
            if not self.autostart:
                raise
            self._start_bridge_process()
            self._wait_for_bridge()

        if self._stream is None:
            raise BridgeUnavailable("Could not establish a connection to the Bridge")
        return self._stream

    def _open_socket(self) -> None:
        connection = socket.create_connection((self.host, self.port), timeout=5)
        connection.settimeout(self.call_timeout)
        self._socket = connection
        self._stream = connection.makefile("rwb")

    def _disconnect(self) -> None:
        for resource in (self._stream, self._socket):
            try:
                if resource is not None:
                    resource.close()
            except OSError:
                pass
        self._stream = None
        self._socket = None

    def _start_bridge_process(self) -> None:
        """Starts the Bridge in the background - Claude Desktop only runs the MCP server."""
        if self._process is not None and self._process.poll() is None:
            return

        command = [
            sys.executable,
            "-m",
            "bridge.main",
            "--host",
            self.host,
            "--port",
            str(self.port),
        ]
        creation_flags = CREATE_NO_WINDOW | DETACHED_PROCESS if os.name == "nt" else 0

        logger.info("Uruchamiam Bridge: %s", " ".join(command))
        self._process = subprocess.Popen(
            command,
            cwd=str(ROOT),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creation_flags,
        )

    def _wait_for_bridge(self) -> None:
        deadline = time.monotonic() + BRIDGE_START_TIMEOUT
        while time.monotonic() < deadline:
            try:
                self._open_socket()
                return
            except OSError:
                time.sleep(0.3)

        raise BridgeUnavailable(
            f"The Bridge did not start within {BRIDGE_START_TIMEOUT:.0f}s - "
            f"uruchom go recznie: python -m bridge.main --port {self.port}"
        )


client = BridgeClient()
server = _McpServer(
    "office-mcp",
    instructions=(
        "Drive open Microsoft Office 2019 apps on Windows through COM. "
        "ppt_* tools handle PowerPoint, xl_* Excel, doc_* Word. "
        "Apps start lazily and changes show up live in the Office window. "
        "PowerPoint coordinates and sizes are in points (1 cm = 28.35 pt), "
        "slide and paragraph indexes count from 1."
    ),
)


def call_bridge(
    app: str,
    action: str,
    params: dict[str, Any] | None = None,
    keep_none: tuple[str, ...] = (),
) -> dict[str, Any]:
    """Sends an action to the Bridge and returns a reply in a fixed JSON shape."""
    cleaned = {
        key: value
        for key, value in (params or {}).items()
        if value is not None or key in keep_none
    }

    try:
        response = client.call(app, action, cleaned)
    except BridgeUnavailable as exc:
        return {
            "ok": False,
            "error": {"type": "BridgeUnavailable", "message": str(exc)},
        }
    except Exception as exc:  # noqa: BLE001 - an MCP tool never raises
        logger.exception("Error calling %s.%s", app, action)
        return {"ok": False, "error": {"type": type(exc).__name__, "message": str(exc)}}

    if response.ok:
        return {"ok": True, "result": response.result}
    return {"ok": False, "error": response.error}


@server.tool()
def office_status() -> dict[str, Any]:
    """COM bridge state: whether the Bridge runs and which Office apps are connected."""
    apps = {}
    for app in ("powerpoint", "excel", "word"):
        apps[app] = call_bridge(app, "status")
    return {"ok": True, "result": {"bridge": client.status(), "apps": apps}}


@server.tool()
def ppt_create_presentation(path: str, template: str | None = None) -> dict[str, Any]:
    """Creates a presentation and saves it at the given path (.pptx).

    The optional 'template' is a path to a .potx/.thmx file with a theme.
    """
    return call_bridge(
        "powerpoint", "create_presentation", {"path": path, "template": template}
    )


@server.tool()
def ppt_open_presentation(path: str) -> dict[str, Any]:
    """Opens an existing presentation; if already open, activates its window."""
    return call_bridge("powerpoint", "open_presentation", {"path": path})


@server.tool()
def ppt_save(path: str | None = None) -> dict[str, Any]:
    """Saves the current presentation; with 'path' saves it as a new file."""
    return call_bridge("powerpoint", "save", {"path": path})


@server.tool()
def ppt_close(save: bool = True) -> dict[str, Any]:
    """Closes the current presentation, saving changes by default."""
    return call_bridge("powerpoint", "close", {"save": save})


@server.tool()
def ppt_get_presentation_info() -> dict[str, Any]:
    """Returns slide count, slide size, theme name and file path."""
    return call_bridge("powerpoint", "get_presentation_info")


@server.tool()
def ppt_get_slide_content(slide_index: int) -> dict[str, Any]:
    """Returns full slide content: shapes, their positions, text and notes."""
    return call_bridge("powerpoint", "get_slide_content", {"slide_index": slide_index})


@server.tool()
def ppt_list_slides() -> dict[str, Any]:
    """Slide list with titles, layout names and shape counts."""
    return call_bridge("powerpoint", "list_slides")


@server.tool()
def ppt_add_slide(
    layout: str = "title_content",
    index: int | None = None,
    title: str | None = None,
) -> dict[str, Any]:
    """Adds a slide with the chosen layout.

    Uklady: title, title_content, two_content, title_only, blank,
    section_header, comparison, picture_with_caption, chart, table.
    Without 'index' the slide goes to the end of the presentation.
    """
    return call_bridge(
        "powerpoint", "add_slide", {"layout": layout, "index": index, "title": title}
    )


@server.tool()
def ppt_delete_slide(slide_index: int) -> dict[str, Any]:
    """Deletes the slide with the given number (counting from 1)."""
    return call_bridge("powerpoint", "delete_slide", {"slide_index": slide_index})


@server.tool()
def ppt_duplicate_slide(slide_index: int) -> dict[str, Any]:
    """Duplicates a slide - the copy lands right after the original."""
    return call_bridge("powerpoint", "duplicate_slide", {"slide_index": slide_index})


@server.tool()
def ppt_reorder_slide(from_index: int, to_index: int) -> dict[str, Any]:
    """Moves a slide to another position in the presentation."""
    return call_bridge(
        "powerpoint", "reorder_slide", {"from_index": from_index, "to_index": to_index}
    )


@server.tool()
def ppt_set_title(slide_index: int, text: str) -> dict[str, Any]:
    """Sets the slide title; if the layout has none, inserts a text box."""
    return call_bridge(
        "powerpoint", "set_title", {"slide_index": slide_index, "text": text}
    )


@server.tool()
def ppt_add_textbox(
    slide_index: int,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    font_size: float | None = None,
    bold: bool = False,
    color: str | None = None,
    align: str | None = None,
) -> dict[str, Any]:
    """Inserts a text box. Coordinates in points, a 16:9 slide is 960x540 pt.

    'color' takes '#RRGGBB' or a colour name, 'align' is left/center/right/justify.
    """
    return call_bridge(
        "powerpoint",
        "add_textbox",
        {
            "slide_index": slide_index,
            "text": text,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "font_size": font_size,
            "bold": bold,
            "color": color,
            "align": align,
        },
    )


@server.tool()
def ppt_add_bullet_list(
    slide_index: int,
    items: list[Any],
    placeholder: str = "content",
) -> dict[str, Any]:
    """Fills a placeholder with a bulleted list.

    'items' to teksty albo obiekty z poziomem wciecia:
    ["Main point", {"text": "Sub point", "level": 2}].
    'placeholder' is "content", "title" or a shape id.
    """
    return call_bridge(
        "powerpoint",
        "add_bullet_list",
        {"slide_index": slide_index, "items": items, "placeholder": placeholder},
    )


@server.tool()
def ppt_find_replace_text(
    old_text: str,
    new_text: str,
    slide_index: int | None = None,
    match_case: bool = False,
) -> dict[str, Any]:
    """Replaces text in the presentation (including tables and shape groups).

    Without 'slide_index' it searches every slide.
    """
    return call_bridge(
        "powerpoint",
        "find_replace_text",
        {
            "old_text": old_text,
            "new_text": new_text,
            "slide_index": slide_index,
            "match_case": match_case,
        },
    )


@server.tool()
def ppt_set_speaker_notes(slide_index: int, text: str) -> dict[str, Any]:
    """Sets the speaker notes for the given slide."""
    return call_bridge(
        "powerpoint", "set_speaker_notes", {"slide_index": slide_index, "text": text}
    )


@server.tool()
def ppt_set_text_style(
    slide_index: int,
    shape_id: int | str,
    font_name: str | None = None,
    font_size: float | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
) -> dict[str, Any]:
    """Formats shape text. Get 'shape_id' from ppt_get_slide_content."""
    return call_bridge(
        "powerpoint",
        "set_text_style",
        {
            "slide_index": slide_index,
            "shape_id": shape_id,
            "font_name": font_name,
            "font_size": font_size,
            "color": color,
            "bold": bold,
            "italic": italic,
            "underline": underline,
        },
    )


@server.tool()
def ppt_apply_theme(theme_name_or_path: str) -> dict[str, Any]:
    """Applies a theme from a .thmx/.potx file or from the Office theme gallery."""
    return call_bridge(
        "powerpoint", "apply_theme", {"theme_name_or_path": theme_name_or_path}
    )


@server.tool()
def ppt_set_background(
    slide_index: int,
    color: str | None = None,
    image_path: str | None = None,
) -> dict[str, Any]:
    """Sets the slide background - a solid colour ('#RRGGBB') or an image file."""
    return call_bridge(
        "powerpoint",
        "set_background",
        {"slide_index": slide_index, "color": color, "image_path": image_path},
    )


@server.tool()
def ppt_set_slide_layout(slide_index: int, layout_name: str) -> dict[str, Any]:
    """Changes the slide layout - by master layout name or a standard name."""
    return call_bridge(
        "powerpoint",
        "set_slide_layout",
        {"slide_index": slide_index, "layout_name": layout_name},
    )


@server.tool()
def ppt_add_image(
    slide_index: int,
    image_path: str,
    left: float,
    top: float,
    width: float | None = None,
    height: float | None = None,
) -> dict[str, Any]:
    """Inserts an image on the slide. Without width/height keeps the proportions."""
    return call_bridge(
        "powerpoint",
        "add_image",
        {
            "slide_index": slide_index,
            "image_path": image_path,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        },
    )


@server.tool()
def ppt_add_chart(
    slide_index: int,
    chart_type: str,
    categories: list[str],
    series_data: dict[str, list[float]] | list[Any],
    left: float,
    top: float,
    width: float,
    height: float,
    title: str | None = None,
) -> dict[str, Any]:
    """Inserts a chart with data.

    'chart_type': bar, column, line, pie, area, scatter, doughnut, radar.
    'series_data': {"Results 2024": [10, 20, 30]} or a list of series
    [{"name": "Wyniki", "values": [10, 20]}].
    """
    return call_bridge(
        "powerpoint",
        "add_chart",
        {
            "slide_index": slide_index,
            "chart_type": chart_type,
            "categories": categories,
            "series_data": series_data,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "title": title,
        },
    )


@server.tool()
def ppt_add_table(
    slide_index: int,
    rows: int,
    cols: int,
    data: list[list[Any]] | None,
    left: float,
    top: float,
    width: float,
    height: float,
) -> dict[str, Any]:
    """Inserts a table and fills it with data (the first row is bolded)."""
    return call_bridge(
        "powerpoint",
        "add_table",
        {
            "slide_index": slide_index,
            "rows": rows,
            "cols": cols,
            "data": data,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        },
    )


@server.tool()
def ppt_add_shape(
    slide_index: int,
    shape_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: str | None = None,
    text: str | None = None,
    line_color: str | None = None,
    line_width: float | None = None,
) -> dict[str, Any]:
    """Inserts a shape: rectangle, rounded_rectangle, oval, triangle, diamond,
    star, arrow_right, callout, cloud, hexagon. 'fill_color'/'line_color'
    take "none" to switch the fill or the outline off."""
    return call_bridge(
        "powerpoint",
        "add_shape",
        {
            "slide_index": slide_index,
            "shape_type": shape_type,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "fill_color": fill_color,
            "text": text,
            "line_color": line_color,
            "line_width": line_width,
        },
    )


@server.tool()
def ppt_group_shapes(
    slide_index: int, shape_ids: list[Any], name: str | None = None
) -> dict[str, Any]:
    """Groups shapes - from now on they move, scale and animate as one.
    Needs at least two shapes."""
    return call_bridge(
        "powerpoint",
        "group_shapes",
        {"slide_index": slide_index, "shape_ids": shape_ids, "name": name},
    )


@server.tool()
def ppt_ungroup_shapes(slide_index: int, shape_id: Any) -> dict[str, Any]:
    """Breaks a group into individual shapes and returns their ids."""
    return call_bridge(
        "powerpoint",
        "ungroup_shapes",
        {"slide_index": slide_index, "shape_id": shape_id},
    )


@server.tool()
def ppt_align_shapes(
    slide_index: int,
    shape_ids: list[Any],
    align: str,
    relative_to_slide: bool = False,
) -> dict[str, Any]:
    """Aligns shapes: left, center, right, top, middle, bottom.
    'relative_to_slide=True' aligns to the slide edges instead of to each other."""
    return call_bridge(
        "powerpoint",
        "align_shapes",
        {
            "slide_index": slide_index,
            "shape_ids": shape_ids,
            "align": align,
            "relative_to_slide": relative_to_slide,
        },
    )


@server.tool()
def ppt_distribute_shapes(
    slide_index: int,
    shape_ids: list[Any],
    direction: str = "horizontal",
    relative_to_slide: bool = False,
) -> dict[str, Any]:
    """Spreads at least three shapes at equal intervals - horizontal
    albo vertical."""
    return call_bridge(
        "powerpoint",
        "distribute_shapes",
        {
            "slide_index": slide_index,
            "shape_ids": shape_ids,
            "direction": direction,
            "relative_to_slide": relative_to_slide,
        },
    )


@server.tool()
def ppt_add_hyperlink(
    slide_index: int,
    shape_id: Any,
    url: str | None = None,
    target_slide: int | None = None,
    tooltip: str | None = None,
) -> dict[str, Any]:
    """Attaches a link to a shape: an external address ('url') or a jump to a
    slide in this deck ('target_slide'). 'tooltip' is the hover hint."""
    return call_bridge(
        "powerpoint",
        "add_hyperlink",
        {
            "slide_index": slide_index,
            "shape_id": shape_id,
            "url": url,
            "target_slide": target_slide,
            "tooltip": tooltip,
        },
    )


@server.tool()
def ppt_set_headers_footers(
    slide_index: int | None = None,
    footer_text: str | None = None,
    show_footer: bool | None = None,
    show_slide_number: bool | None = None,
    show_date: bool | None = None,
) -> dict[str, Any]:
    """Footer, slide number and date. Without 'slide_index' covers every slide.
    Sam 'footer_text' automatycznie wlacza widocznosc stopki."""
    return call_bridge(
        "powerpoint",
        "set_headers_footers",
        {
            "slide_index": slide_index,
            "footer_text": footer_text,
            "show_footer": show_footer,
            "show_slide_number": show_slide_number,
            "show_date": show_date,
        },
    )


@server.tool()
def ppt_add_media(
    slide_index: int,
    media_path: str,
    left: float,
    top: float,
    width: float | None = None,
    height: float | None = None,
    autoplay: bool = False,
) -> dict[str, Any]:
    """Inserts video or audio embedded in the presentation. 'autoplay=True'
    attaches a play effect starting with the previous one instead of on click."""
    return call_bridge(
        "powerpoint",
        "add_media",
        {
            "slide_index": slide_index,
            "media_path": media_path,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "autoplay": autoplay,
        },
    )


@server.tool()
def ppt_list_smartart_layouts(
    search: str | None = None, category: str | None = None
) -> dict[str, Any]:
    """SmartArt layouts: key, name and category. NOTE: 'name' is localised
    (a non-English Office returns translated names), so pick a layout by
    'key' - it is identical in every language version. 'category' is not
    translated either: list, process, cycle,
    hierarchy, relationship, matrix, pyramid, picture."""
    return call_bridge(
        "powerpoint",
        "list_smartart_layouts",
        {"search": search, "category": category},
    )


@server.tool()
def ppt_add_smartart(
    slide_index: int,
    layout: Any,
    items: list[Any],
    left: float,
    top: float,
    width: float,
    height: float,
) -> dict[str, Any]:
    """Inserts a SmartArt diagram and fills it with text. 'layout' is a key
    ('bProcess3', 'hierarchy1'), a number or a name from ppt_list_smartart_layouts -
    the key is safer, because names are translated into the Office language.
    'items' przyjmuje teksty albo slowniki {"text": ..., "level": 2} -
    level 2+ creates child nodes."""
    return call_bridge(
        "powerpoint",
        "add_smartart",
        {
            "slide_index": slide_index,
            "layout": layout,
            "items": items,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        },
    )


@server.tool()
def ppt_list_sections() -> dict[str, Any]:
    """Presentation sections with their first slide and slide count."""
    return call_bridge("powerpoint", "list_sections", {})


@server.tool()
def ppt_add_section(name: str, before_slide: int = 1) -> dict[str, Any]:
    """Creates a section starting at the given slide."""
    return call_bridge(
        "powerpoint", "add_section", {"name": name, "before_slide": before_slide}
    )


@server.tool()
def ppt_delete_section(
    section_index: int, delete_slides: bool = False
) -> dict[str, Any]:
    """Deletes a section; 'delete_slides=True' also removes its slides."""
    return call_bridge(
        "powerpoint",
        "delete_section",
        {"section_index": section_index, "delete_slides": delete_slides},
    )


@server.tool()
def ppt_slideshow(
    command: str = "start", slide_index: int | None = None
) -> dict[str, Any]:
    """Controls the slide show: 'start' (optionally from 'slide_index'), 'stop',
    'goto' (requires 'slide_index')."""
    return call_bridge(
        "powerpoint",
        "slideshow",
        {"command": command, "slide_index": slide_index},
    )


@server.tool()
def ppt_copy_slide_to(
    slide_index: int, target_path: str, position: int | None = None
) -> dict[str, Any]:
    """Copies a slide into another, existing presentation. Without 'position'
    the slide goes to the end. Does not use the clipboard."""
    return call_bridge(
        "powerpoint",
        "copy_slide_to",
        {
            "slide_index": slide_index,
            "target_path": target_path,
            "position": position,
        },
    )


@server.tool()
def ppt_get_theme() -> dict[str, Any]:
    """Returns the presentation theme's colour palette and fonts."""
    return call_bridge("powerpoint", "get_theme", {})


@server.tool()
def ppt_set_theme_colors(colors: dict[str, str]) -> dict[str, Any]:
    """Podmienia kolory w palecie motywu, np.
    {"accent1": "#10A37F", "dark1": "#0B1014", "light1": "#ECF2F0"}.
    Names: dark1/text1, light1/background1, dark2, light2, accent1..accent6,
    hyperlink, followed_hyperlink. Set once, the theme applies to every
    slide - no need to repeat the colour on every shape."""
    return call_bridge("powerpoint", "set_theme_colors", {"colors": colors})


@server.tool()
def ppt_set_theme_fonts(
    major: str | None = None, minor: str | None = None
) -> dict[str, Any]:
    """Sets theme fonts: 'major' for headings, 'minor' for body text."""
    return call_bridge(
        "powerpoint", "set_theme_fonts", {"major": major, "minor": minor}
    )


@server.tool()
def ppt_set_master_background(
    color: str | None = None,
    image_path: str | None = None,
    apply_to_slides: bool = True,
) -> dict[str, Any]:
    """Sets the background on the slide master - once for the whole deck, instead
    of calling ppt_set_background for every slide separately."""
    return call_bridge(
        "powerpoint",
        "set_master_background",
        {
            "color": color,
            "image_path": image_path,
            "apply_to_slides": apply_to_slides,
        },
    )


@server.tool()
def ppt_set_shape_format(
    slide_index: int,
    shape_id: Any,
    fill_color: str | None = None,
    fill_transparency: float | None = None,
    gradient_from: str | None = None,
    gradient_to: str | None = None,
    gradient_style: str = "vertical",
    line_color: str | None = None,
    line_width: float | None = None,
    line_dash: str | None = None,
    shadow: bool | None = None,
    shadow_color: str | None = None,
    shadow_blur: float | None = None,
    shadow_offset_x: float | None = None,
    shadow_offset_y: float | None = None,
    shadow_transparency: float | None = None,
    corner_radius: float | None = None,
) -> dict[str, Any]:
    """Look of an existing shape. 'gradient_from' + 'gradient_to' turn on a
    gradient dwukolorowy ('gradient_style': horizontal, vertical, diagonal_up,
    diagonal_down, from_corner, from_center). Przezroczystosci 0.0-1.0.
    'line_dash': solid, dash, round_dot, long_dash i pokrewne.
    'corner_radius' 0.0-0.5 dziala na rounded_rectangle."""
    return call_bridge(
        "powerpoint",
        "set_shape_format",
        {
            "slide_index": slide_index,
            "shape_id": shape_id,
            "fill_color": fill_color,
            "fill_transparency": fill_transparency,
            "gradient_from": gradient_from,
            "gradient_to": gradient_to,
            "gradient_style": gradient_style,
            "line_color": line_color,
            "line_width": line_width,
            "line_dash": line_dash,
            "shadow": shadow,
            "shadow_color": shadow_color,
            "shadow_blur": shadow_blur,
            "shadow_offset_x": shadow_offset_x,
            "shadow_offset_y": shadow_offset_y,
            "shadow_transparency": shadow_transparency,
            "corner_radius": corner_radius,
        },
    )


@server.tool()
def ppt_set_paragraph_format(
    slide_index: int,
    shape_id: Any,
    paragraph: int | None = None,
    line_spacing: float | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    alignment: str | None = None,
    vertical_anchor: str | None = None,
    autosize: bool | None = None,
    word_wrap: bool | None = None,
    margin: float | None = None,
) -> dict[str, Any]:
    """Paragraph typography: line spacing (a multiple, 1.0 = single), gaps
    przed/po w punktach, wyrownanie (left/center/right/justify), kotwica pionowa
    (top/middle/bottom), autodopasowanie ramki i marginesy wewnetrzne.
    Without 'paragraph' the change covers all the text of the shape."""
    return call_bridge(
        "powerpoint",
        "set_paragraph_format",
        {
            "slide_index": slide_index,
            "shape_id": shape_id,
            "paragraph": paragraph,
            "line_spacing": line_spacing,
            "space_before": space_before,
            "space_after": space_after,
            "alignment": alignment,
            "vertical_anchor": vertical_anchor,
            "autosize": autosize,
            "word_wrap": word_wrap,
            "margin": margin,
        },
    )


@server.tool()
def ppt_format_chart(
    slide_index: int,
    shape_id: Any,
    series_colors: list[str] | None = None,
    text_color: str | None = None,
    background: str | None = None,
    legend: Any = None,
    data_labels: bool | None = None,
    gridlines: bool | None = None,
    title: str | None = None,
    value_axis_min: float | None = None,
    value_axis_max: float | None = None,
) -> dict[str, Any]:
    """Tunes a chart to the slide colours: series colours, axis and legend text
    i legendy, tlo ('none' = przezroczyste), pozycja legendy
    (bottom/top/left/right albo False), etykiety danych, linie siatki, tytul."""
    return call_bridge(
        "powerpoint",
        "format_chart",
        {
            "slide_index": slide_index,
            "shape_id": shape_id,
            "series_colors": series_colors,
            "text_color": text_color,
            "background": background,
            "legend": legend,
            "data_labels": data_labels,
            "gridlines": gridlines,
            "title": title,
            "value_axis_min": value_axis_min,
            "value_axis_max": value_axis_max,
        },
    )


@server.tool()
def ppt_delete_shape(slide_index: int, shape_id: Any) -> dict[str, Any]:
    """Deletes a shape from the slide; 'shape_id' is an id, a shape name or the
    'title'/'content'."""
    return call_bridge(
        "powerpoint",
        "delete_shape",
        {"slide_index": slide_index, "shape_id": shape_id},
    )


@server.tool()
def ppt_set_shape_position(
    slide_index: int,
    shape_id: Any,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    rotation: float | None = None,
) -> dict[str, Any]:
    """Moves, scales and rotates an existing shape. Coordinates in points
    (a 16:9 slide is 960 x 540 pt), 'rotation' in degrees. Pass only the
    fields you want to change."""
    return call_bridge(
        "powerpoint",
        "set_shape_position",
        {
            "slide_index": slide_index,
            "shape_id": shape_id,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "rotation": rotation,
        },
    )


@server.tool()
def ppt_set_shape_order(
    slide_index: int, shape_id: Any, order: str = "front"
) -> dict[str, Any]:
    """Changes the shape layer: front (to the top), back (to the bottom),
    forward (krok w gore), backward (krok w dol)."""
    return call_bridge(
        "powerpoint",
        "set_shape_order",
        {"slide_index": slide_index, "shape_id": shape_id, "order": order},
    )


@server.tool()
def ppt_export_slide(
    slide_index: int,
    path: str,
    width: int | None = None,
    height: int | None = None,
) -> dict[str, Any]:
    """Saves the slide as an image - the format follows the file extension
    (.png, .jpg, .gif, .bmp, .wmf, .emf). Without dimensions the image is
    1920 px szerokosci. Sluzy do obejrzenia efektu i poprawienia ukladu."""
    return call_bridge(
        "powerpoint",
        "export_slide",
        {"slide_index": slide_index, "path": path, "width": width, "height": height},
    )


@server.tool()
def ppt_export_pdf(path: str, embed_fonts: bool = True) -> dict[str, Any]:
    """Exports the whole presentation to PDF; does not change the file open
    w PowerPoincie."""
    return call_bridge(
        "powerpoint", "export_pdf", {"path": path, "embed_fonts": embed_fonts}
    )


@server.tool()
def ppt_add_animation(
    slide_index: int,
    shape_id: Any,
    effect: str = "fade",
    trigger: str = "after_previous",
    level: str = "shape",
    duration: float | None = None,
    delay: float | None = None,
    exit_effect: bool = False,
) -> dict[str, Any]:
    """Animates a shape on a slide. 'shape_id' is an id, a shape name or the
    'title'/'content'. Efekty: fade, fly, wipe, zoom, float, grow_and_turn,
    rise_up, split, wheel, bounce, spin, grow_shrink, teeter i inne.
    Wyzwalacze: on_click, with_previous, after_previous. 'level' = shape albo
    by_paragraph (text paragraph by paragraph). 'duration' and 'delay' in seconds."""
    return call_bridge(
        "powerpoint",
        "add_animation",
        {
            "slide_index": slide_index,
            "shape_id": shape_id,
            "effect": effect,
            "trigger": trigger,
            "level": level,
            "duration": duration,
            "delay": delay,
            "exit_effect": exit_effect,
        },
    )


@server.tool()
def ppt_list_animations(slide_index: int) -> dict[str, Any]:
    """Returns the slide's animations in playback order plus its transition."""
    return call_bridge("powerpoint", "list_animations", {"slide_index": slide_index})


@server.tool()
def ppt_set_transition(
    effect: str = "fade",
    slide_index: int | None = None,
    duration: float | None = None,
    advance_on_click: bool = True,
    advance_after: float | None = None,
) -> dict[str, Any]:
    """Sets the transition between slides; without 'slide_index' covers the whole
    prezentacje. Efekty: fade, fade_smoothly, push_left, wipe_right, cover_up,
    split_vertical_out, zoom_in, morph, honeycomb, gallery_left, cube_left,
    doors_vertical, curtains, prestige i inne. 'duration' w sekundach,
    'advance_after' turns on automatic slide change after the given time."""
    return call_bridge(
        "powerpoint",
        "set_transition",
        {
            "effect": effect,
            "slide_index": slide_index,
            "duration": duration,
            "advance_on_click": advance_on_click,
            "advance_after": advance_after,
        },
    )


@server.tool()
def xl_create_workbook(path: str) -> dict[str, Any]:
    """Creates a new workbook and saves it at the given path (.xlsx)."""
    return call_bridge("excel", "create_workbook", {"path": path})


@server.tool()
def xl_open_workbook(path: str) -> dict[str, Any]:
    """Opens an existing workbook; if already open, activates its window."""
    return call_bridge("excel", "open_workbook", {"path": path})


@server.tool()
def xl_save(path: str | None = None) -> dict[str, Any]:
    """Saves the current workbook; with 'path' saves it as a new file."""
    return call_bridge("excel", "save", {"path": path})


@server.tool()
def xl_close(save: bool = True) -> dict[str, Any]:
    """Closes the current workbook, saving changes by default."""
    return call_bridge("excel", "close", {"save": save})


@server.tool()
def xl_add_sheet(name: str, index: int | None = None) -> dict[str, Any]:
    """Adds a sheet with the given name; without 'index' it goes to the end."""
    return call_bridge("excel", "add_sheet", {"name": name, "index": index})


@server.tool()
def xl_delete_sheet(name: str) -> dict[str, Any]:
    """Deletes the sheet with the given name."""
    return call_bridge("excel", "delete_sheet", {"name": name})


@server.tool()
def xl_rename_sheet(old_name: str, new_name: str) -> dict[str, Any]:
    """Renames a sheet."""
    return call_bridge(
        "excel", "rename_sheet", {"old_name": old_name, "new_name": new_name}
    )


@server.tool()
def xl_get_workbook_info() -> dict[str, Any]:
    """Returns the sheet list with their data ranges, the active sheet and file path."""
    return call_bridge("excel", "get_workbook_info")


@server.tool()
def xl_get_range_values(sheet: str, range_ref: str) -> dict[str, Any]:
    """Reads range values (e.g. "A1:D10") as a 2D array."""
    return call_bridge("excel", "get_range_values", {"sheet": sheet, "range_ref": range_ref})


@server.tool()
def xl_get_used_range(sheet: str) -> dict[str, Any]:
    """Returns the actually filled area of the sheet along with its data."""
    return call_bridge("excel", "get_used_range", {"sheet": sheet})


@server.tool()
def xl_set_cell(sheet: str, cell_ref: str, value: Any) -> dict[str, Any]:
    """Writes a value into a cell (e.g. sheet="Budget", cell_ref="B2")."""
    return call_bridge(
        "excel",
        "set_cell",
        {"sheet": sheet, "cell_ref": cell_ref, "value": value},
        keep_none=("value",),
    )


@server.tool()
def xl_set_range(sheet: str, start_cell: str, values_2d: list[list[Any]]) -> dict[str, Any]:
    """Pastes a whole matrix at once, starting at 'start_cell'.

    Much faster than writing cell by cell.
    """
    return call_bridge(
        "excel",
        "set_range",
        {"sheet": sheet, "start_cell": start_cell, "values_2d": values_2d},
    )


@server.tool()
def xl_set_formula(sheet: str, cell_ref: str, formula: str) -> dict[str, Any]:
    """Writes a formula (e.g. "=SUM(A1:A10)") and returns the computed result."""
    return call_bridge(
        "excel", "set_formula", {"sheet": sheet, "cell_ref": cell_ref, "formula": formula}
    )


@server.tool()
def xl_clear_range(sheet: str, range_ref: str, contents_only: bool = True) -> dict[str, Any]:
    """Clears a range - values only by default, optionally formatting too."""
    return call_bridge(
        "excel",
        "clear_range",
        {"sheet": sheet, "range_ref": range_ref, "contents_only": contents_only},
    )


@server.tool()
def xl_insert_rows(sheet: str, start_row: int, count: int = 1) -> dict[str, Any]:
    """Inserts rows, pushing existing ones down."""
    return call_bridge(
        "excel", "insert_rows", {"sheet": sheet, "start_row": start_row, "count": count}
    )


@server.tool()
def xl_delete_rows(sheet: str, start_row: int, count: int = 1) -> dict[str, Any]:
    """Deletes rows, pulling the rest up."""
    return call_bridge(
        "excel", "delete_rows", {"sheet": sheet, "start_row": start_row, "count": count}
    )


@server.tool()
def xl_insert_columns(sheet: str, start_col: str | int, count: int = 1) -> dict[str, Any]:
    """Inserts columns; 'start_col' takes a letter ("C") or a number (3)."""
    return call_bridge(
        "excel",
        "insert_columns",
        {"sheet": sheet, "start_col": start_col, "count": count},
    )


@server.tool()
def xl_delete_columns(sheet: str, start_col: Any, count: int = 1) -> dict[str, Any]:
    """Deletes columns; 'start_col' takes a letter ("C") or a number (3)."""
    return call_bridge(
        "excel",
        "delete_columns",
        {"sheet": sheet, "start_col": start_col, "count": count},
    )


@server.tool()
def xl_set_row_height(sheet: str, row: int, height: Any) -> dict[str, Any]:
    """Row height in points; height="auto" fits it to the content."""
    return call_bridge(
        "excel", "set_row_height", {"sheet": sheet, "row": row, "height": height}
    )


@server.tool()
def xl_find_replace(
    old_text: str,
    new_text: str,
    sheet: str | None = None,
    range_ref: str | None = None,
    match_case: bool = False,
    whole_cell: bool = False,
) -> dict[str, Any]:
    """Replaces text; without 'sheet' it walks every sheet.
    'whole_cell=True' requires the whole cell content to equal the search text."""
    return call_bridge(
        "excel",
        "find_replace",
        {
            "old_text": old_text,
            "new_text": new_text,
            "sheet": sheet,
            "range_ref": range_ref,
            "match_case": match_case,
            "whole_cell": whole_cell,
        },
    )


@server.tool()
def xl_sort_range(
    sheet: str,
    range_ref: str,
    sort_by: Any,
    order: str = "ascending",
    has_headers: bool = True,
) -> dict[str, Any]:
    """Sorts a range by column 'sort_by' (letter, number or cell address).
    'order' to ascending albo descending."""
    return call_bridge(
        "excel",
        "sort_range",
        {
            "sheet": sheet,
            "range_ref": range_ref,
            "sort_by": sort_by,
            "order": order,
            "has_headers": has_headers,
        },
    )


@server.tool()
def xl_set_autofilter(
    sheet: str, range_ref: str | None = None, enable: bool = True
) -> dict[str, Any]:
    """Turns AutoFilter on or off; without 'range_ref' covers the used range."""
    return call_bridge(
        "excel",
        "set_autofilter",
        {"sheet": sheet, "range_ref": range_ref, "enable": enable},
    )


@server.tool()
def xl_copy_range(
    sheet: str,
    range_ref: str,
    target_cell: str,
    target_sheet: str | None = None,
    paste: str = "all",
) -> dict[str, Any]:
    """Copies a range to the same or another place. 'paste' is all, values
    (results only, no formulas) or formats."""
    return call_bridge(
        "excel",
        "copy_range",
        {
            "sheet": sheet,
            "range_ref": range_ref,
            "target_cell": target_cell,
            "target_sheet": target_sheet,
            "paste": paste,
        },
    )


@server.tool()
def xl_add_data_validation(
    sheet: str,
    range_ref: str,
    validation_type: str = "list",
    values: Any = None,
    formula: str | None = None,
    formula2: str | None = None,
    operator: str | None = None,
    alert: str = "stop",
    input_message: str | None = None,
    error_message: str | None = None,
) -> dict[str, Any]:
    """Data validation. For a dropdown list just pass 'values' (a list of
    entries or a range reference). The other types: whole_number,
    decimal, date, time, text_length, custom - uzywaja 'formula' i 'operator'."""
    return call_bridge(
        "excel",
        "add_data_validation",
        {
            "sheet": sheet,
            "range_ref": range_ref,
            "validation_type": validation_type,
            "values": values,
            "formula": formula,
            "formula2": formula2,
            "operator": operator,
            "alert": alert,
            "input_message": input_message,
            "error_message": error_message,
        },
    )


@server.tool()
def xl_get_cell_formula(sheet: str, range_ref: str) -> dict[str, Any]:
    """Returns range formulas (not computed values) along with the results."""
    return call_bridge(
        "excel", "get_cell_formula", {"sheet": sheet, "range_ref": range_ref}
    )


@server.tool()
def xl_export_pdf(
    path: str, sheet: str | None = None, range_ref: str | None = None
) -> dict[str, Any]:
    """Exports the workbook, a single sheet or a range to PDF."""
    return call_bridge(
        "excel",
        "export_pdf",
        {"path": path, "sheet": sheet, "range_ref": range_ref},
    )


@server.tool()
def xl_export_range_image(sheet: str, range_ref: str, path: str) -> dict[str, Any]:
    """Saves a range as an image (.png/.jpg/.gif). Use it to look at the
    formatting and fix it - the same idea as ppt_export_slide for slides."""
    return call_bridge(
        "excel",
        "export_range_image",
        {"sheet": sheet, "range_ref": range_ref, "path": path},
    )


@server.tool()
def xl_format_chart(
    sheet: str,
    chart: Any = 1,
    series_colors: list[str] | None = None,
    text_color: str | None = None,
    background: str | None = None,
    legend: Any = None,
    data_labels: bool | None = None,
    gridlines: bool | None = None,
    title: str | None = None,
    value_axis_min: float | None = None,
    value_axis_max: float | None = None,
) -> dict[str, Any]:
    """Tunes a chart in the sheet: series colours, axis and legend text colour,
    tlo ('none' = przezroczyste), pozycja legendy, etykiety, siatka, tytul.
    'chart' is the number or name of the chart object in the sheet."""
    return call_bridge(
        "excel",
        "format_chart",
        {
            "sheet": sheet,
            "chart": chart,
            "series_colors": series_colors,
            "text_color": text_color,
            "background": background,
            "legend": legend,
            "data_labels": data_labels,
            "gridlines": gridlines,
            "title": title,
            "value_axis_min": value_axis_min,
            "value_axis_max": value_axis_max,
        },
    )


@server.tool()
def xl_set_cell_format(
    sheet: str,
    range_ref: str,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    fill_color: str | None = None,
    number_format: str | None = None,
    align: str | None = None,
    wrap_text: bool | None = None,
) -> dict[str, Any]:
    """Formats a range of cells.

    'number_format' to maska Excela, np. "0.00", "# ##0 zl", "0%".
    Colours take '#RRGGBB' or a name (red, blue, green...).
    """
    return call_bridge(
        "excel",
        "set_cell_format",
        {
            "sheet": sheet,
            "range_ref": range_ref,
            "bold": bold,
            "italic": italic,
            "font_size": font_size,
            "font_color": font_color,
            "fill_color": fill_color,
            "number_format": number_format,
            "align": align,
            "wrap_text": wrap_text,
        },
    )


@server.tool()
def xl_set_column_width(sheet: str, column: str | int, width: float | str) -> dict[str, Any]:
    """Sets column width; width="auto" fits it to the contents."""
    return call_bridge(
        "excel", "set_column_width", {"sheet": sheet, "column": column, "width": width}
    )


@server.tool()
def xl_merge_cells(sheet: str, range_ref: str, center: bool = True) -> dict[str, Any]:
    """Merges the cells of a range, centring the content by default."""
    return call_bridge(
        "excel", "merge_cells", {"sheet": sheet, "range_ref": range_ref, "center": center}
    )


@server.tool()
def xl_apply_conditional_formatting(
    sheet: str,
    range_ref: str,
    rule_type: str,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Dodaje formatowanie warunkowe.

    rule_type='cell_value'    params={"operator": "greater", "formula1": 1000,
                                      "fill_color": "red"}
    rule_type='expression'    params={"formula": "=$D2>1000", "bold": true}
    rule_type='text_contains' params={"text": "TODO", "fill_color": "yellow"}
    rule_type='color_scale'   params={"colors": ["#F8696B", "#FFEB84", "#63BE7B"]}
    rule_type='data_bar'      params={"color": "#638EC6"}

    Operatory: greater, less, equal, not_equal, greater_equal, less_equal,
    between, not_between.
    """
    return call_bridge(
        "excel",
        "apply_conditional_formatting",
        {
            "sheet": sheet,
            "range_ref": range_ref,
            "rule_type": rule_type,
            "params": params,
        },
    )


@server.tool()
def xl_freeze_panes(sheet: str, cell_ref: str) -> dict[str, Any]:
    """Freezes rows and columns above and to the left of a cell (e.g. "A2")."""
    return call_bridge("excel", "freeze_panes", {"sheet": sheet, "cell_ref": cell_ref})


@server.tool()
def xl_add_chart(
    sheet: str,
    chart_type: str,
    data_range: str,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str | None = None,
) -> dict[str, Any]:
    """Inserts a chart based on a data range from the same sheet.

    'chart_type': column, bar, line, pie, area, scatter, doughnut, radar.
    Position and size in points.
    """
    return call_bridge(
        "excel",
        "add_chart",
        {
            "sheet": sheet,
            "chart_type": chart_type,
            "data_range": data_range,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "title": title,
        },
    )


@server.tool()
def xl_create_table(
    sheet: str,
    range_ref: str,
    table_name: str,
    has_headers: bool = True,
    style: str = "TableStyleMedium2",
) -> dict[str, Any]:
    """Turns a range into a native Excel table (with filters and a style)."""
    return call_bridge(
        "excel",
        "create_table",
        {
            "sheet": sheet,
            "range_ref": range_ref,
            "table_name": table_name,
            "has_headers": has_headers,
            "style": style,
        },
    )


@server.tool()
def xl_add_pivot_table(
    sheet: str,
    source_range: str,
    dest_cell: str,
    rows: list[str] | None = None,
    columns: list[str] | None = None,
    values: list[Any] | None = None,
    dest_sheet: str | None = None,
    table_name: str = "TabelaPrzestawna1",
) -> dict[str, Any]:
    """Builds a pivot table from a source range (the first row is the headers).

    'values' is field names or objects {"field": "Amount", "function": "average"}.
    Funkcje: sum, count, average, max, min, product, count_numbers, std_dev.
    """
    return call_bridge(
        "excel",
        "add_pivot_table",
        {
            "sheet": sheet,
            "source_range": source_range,
            "dest_cell": dest_cell,
            "rows": rows,
            "columns": columns,
            "values": values,
            "dest_sheet": dest_sheet,
            "table_name": table_name,
        },
    )


@server.tool()
def doc_set_paragraph_format(
    paragraph_index: int | None = None,
    count: int = 1,
    style: str | None = None,
    body_text_only: bool = False,
    line_spacing: float | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    first_line_indent: float | None = None,
    left_indent: float | None = None,
    right_indent: float | None = None,
    alignment: str | None = None,
    keep_with_next: bool | None = None,
    page_break_before: bool | None = None,
    widow_control: bool | None = None,
    unit: str = "pt",
) -> dict[str, Any]:
    """Interlinia, wciecia i lamanie akapitow - podstawa skladu pracy dyplomowej.
    Scope: 'style' changes a style definition (e.g. "Normal" = all body text at
    once), 'paragraph_index' with 'count' covers specific paragraphs, and
    neither means every paragraph or, with 'body_text_only', body text only.
    'line_spacing' 1.0 / 1.5 / 2.0 albo dowolna wielokrotnosc. Odstepy i wciecia
    w jednostce 'unit' (pt, cm, mm, in)."""
    return call_bridge(
        "word",
        "set_paragraph_format",
        {
            "paragraph_index": paragraph_index,
            "count": count,
            "style": style,
            "body_text_only": body_text_only,
            "line_spacing": line_spacing,
            "space_before": space_before,
            "space_after": space_after,
            "first_line_indent": first_line_indent,
            "left_indent": left_indent,
            "right_indent": right_indent,
            "alignment": alignment,
            "keep_with_next": keep_with_next,
            "page_break_before": page_break_before,
            "widow_control": widow_control,
            "unit": unit,
        },
    )


@server.tool()
def doc_set_heading_numbering(
    enable: bool = True, levels: int = 3, indent: float = 0.0
) -> dict[str, Any]:
    """Wlacza numeracje rozdzialow 1., 1.1, 1.1.1 powiazana ze stylami naglowkow.
    Only heading paragraphs get numbered; body text is left alone.
    'enable=False' zdejmuje numeracje."""
    return call_bridge(
        "word",
        "set_heading_numbering",
        {"enable": enable, "levels": levels, "indent": indent},
    )


@server.tool()
def doc_add_caption(
    paragraph_index: int,
    text: str,
    label: str = "figure",
    above: bool = False,
) -> dict[str, Any]:
    """A numbered caption next to a paragraph. 'label' is a built-in label
    ('figure', 'table', 'equation') or any custom text - a custom label goes
    into the document verbatim, so a localised document can pass
    label="Figura" or similar. Numbering is a Word field, so later captions
    przenumerowuja wczesniejsze - po zmianach wywolaj doc_update_fields."""
    return call_bridge(
        "word",
        "add_caption",
        {
            "paragraph_index": paragraph_index,
            "text": text,
            "label": label,
            "above": above,
        },
    )


@server.tool()
def doc_insert_table_of_figures(
    label: str = "figure", position: Any = "end"
) -> dict[str, Any]:
    """Spis rysunkow albo tabel zbudowany z podpisow. 'position': start, end albo
    a paragraph number to place the table after."""
    return call_bridge(
        "word",
        "insert_table_of_figures",
        {"label": label, "position": position},
    )


@server.tool()
def doc_update_fields() -> dict[str, Any]:
    """Odswieza pola: spis tresci, spisy rysunkow i numeracje podpisow. Spis tresci
    inserted before the chapters are written stays empty until refreshed."""
    return call_bridge("word", "update_fields", {})


@server.tool()
def doc_set_page_setup(
    orientation: str | None = None,
    gutter: float | None = None,
    mirror_margins: bool | None = None,
    different_first_page: bool | None = None,
    section: int | None = None,
    unit: str = "cm",
) -> dict[str, Any]:
    """Orientacja strony (portrait/landscape), margines na oprawe ('gutter')
    i marginesy lustrzane - ustawienia druku dwustronnego pracy dyplomowej."""
    return call_bridge(
        "word",
        "set_page_setup",
        {
            "orientation": orientation,
            "gutter": gutter,
            "mirror_margins": mirror_margins,
            "different_first_page": different_first_page,
            "section": section,
            "unit": unit,
        },
    )


@server.tool()
def doc_export_pdf(path: str, open_after: bool = False) -> dict[str, Any]:
    """Exports the document to PDF; does not change the file open in Word."""
    return call_bridge(
        "word", "export_pdf", {"path": path, "open_after": open_after}
    )


@server.tool()
def doc_get_paragraph(paragraph_index: int, count: int = 1) -> dict[str, Any]:
    """Reads paragraphs with their style, alignment and outline level."""
    return call_bridge(
        "word",
        "get_paragraph",
        {"paragraph_index": paragraph_index, "count": count},
    )


@server.tool()
def doc_delete_paragraph(paragraph_index: int, count: int = 1) -> dict[str, Any]:
    """Deletes a paragraph, or several in a row, from the given index."""
    return call_bridge(
        "word",
        "delete_paragraph",
        {"paragraph_index": paragraph_index, "count": count},
    )


@server.tool()
def doc_insert_paragraph(
    text: str,
    paragraph_index: int | None = None,
    after: bool = False,
    style: str | None = None,
) -> dict[str, Any]:
    """Inserts a paragraph at a specific place. Without 'paragraph_index' it
    appends at the end; with an index it inserts before that paragraph, or after it with 'after=True'."""
    return call_bridge(
        "word",
        "insert_paragraph",
        {
            "text": text,
            "paragraph_index": paragraph_index,
            "after": after,
            "style": style,
        },
    )


@server.tool()
def doc_add_hyperlink(
    url: str,
    text: str | None = None,
    paragraph_index: int | None = None,
    tooltip: str | None = None,
) -> dict[str, Any]:
    """Inserts a hyperlink; without 'paragraph_index' appends it at the end."""
    return call_bridge(
        "word",
        "add_hyperlink",
        {
            "url": url,
            "text": text,
            "paragraph_index": paragraph_index,
            "tooltip": tooltip,
        },
    )


@server.tool()
def doc_add_footnote(paragraph_index: int, text: str) -> dict[str, Any]:
    """Adds a footnote at the end of the given paragraph."""
    return call_bridge(
        "word",
        "add_footnote",
        {"paragraph_index": paragraph_index, "text": text},
    )


@server.tool()
def doc_insert_section_break(
    break_type: str = "next_page", paragraph_index: int | None = None
) -> dict[str, Any]:
    """Podzial sekcji: next_page, continuous, even_page, odd_page. Sekcje maja
    their own margins, columns, headers and footers."""
    return call_bridge(
        "word",
        "insert_section_break",
        {"break_type": break_type, "paragraph_index": paragraph_index},
    )


@server.tool()
def doc_set_columns(
    count: int = 1, section: int = 1, spacing: float | None = None
) -> dict[str, Any]:
    """Sets the number of text columns in a section (newspaper layout); 'spacing' in points."""
    return call_bridge(
        "word",
        "set_columns",
        {"count": count, "section": section, "spacing": spacing},
    )


@server.tool()
def doc_set_default_font(
    name: str | None = None, size: float | None = None
) -> dict[str, Any]:
    """Changes the Normal style font - the basis of the whole document, instead
    of setting the font paragraph by paragraph."""
    return call_bridge("word", "set_default_font", {"name": name, "size": size})


@server.tool()
def doc_format_table(
    table_index: int = 1,
    style: str | None = None,
    borders: bool | None = None,
    header_bold: bool | None = None,
    header_fill: str | None = None,
    column_widths: list[float] | None = None,
    autofit: bool | None = None,
) -> dict[str, Any]:
    """Formats an inserted table. 'style' takes language-independent names:
    normal, light_shading, light_list, light_grid, medium_shading1,
    medium_grid1..3, dark_list, colorful_shading, colorful_list, colorful_grid
    (plus the _accent1 variants). Column widths in points."""
    return call_bridge(
        "word",
        "format_table",
        {
            "table_index": table_index,
            "style": style,
            "borders": borders,
            "header_bold": header_bold,
            "header_fill": header_fill,
            "column_widths": column_widths,
            "autofit": autofit,
        },
    )


@server.tool()
def doc_create_document(path: str, template: str | None = None) -> dict[str, Any]:
    """Creates a document and saves it at the given path (.docx).

    The optional 'template' is a path to a .dotx file.
    """
    return call_bridge("word", "create_document", {"path": path, "template": template})


@server.tool()
def doc_open_document(path: str) -> dict[str, Any]:
    """Opens an existing document; if already open, activates its window."""
    return call_bridge("word", "open_document", {"path": path})


@server.tool()
def doc_save(path: str | None = None) -> dict[str, Any]:
    """Saves the current document; with 'path' saves it as a new file."""
    return call_bridge("word", "save", {"path": path})


@server.tool()
def doc_close(save: bool = True) -> dict[str, Any]:
    """Closes the current document, saving changes by default."""
    return call_bridge("word", "close", {"save": save})


@server.tool()
def doc_get_document_info() -> dict[str, Any]:
    """Returns page, word and character counts, the template name and file path."""
    return call_bridge("word", "get_document_info")


@server.tool()
def doc_get_full_text() -> dict[str, Any]:
    """Returns the whole document text."""
    return call_bridge("word", "get_full_text")


@server.tool()
def doc_get_outline() -> dict[str, Any]:
    """Returns the heading structure (level, text, paragraph index)."""
    return call_bridge("word", "get_outline")


@server.tool()
def doc_add_paragraph(text: str, style: str | None = None) -> dict[str, Any]:
    """Appends a paragraph at the end of the document.

    'style' accepts English names (Normal, Heading 1, Quote, Caption) even in a
    localised Word.
    """
    return call_bridge("word", "add_paragraph", {"text": text, "style": style})


@server.tool()
def doc_add_heading(text: str, level: int = 1) -> dict[str, Any]:
    """Appends a level 1-9 heading (Heading N style)."""
    return call_bridge("word", "add_heading", {"text": text, "level": level})


@server.tool()
def doc_insert_page_break() -> dict[str, Any]:
    """Inserts a hard page break at the end of the document."""
    return call_bridge("word", "insert_page_break")


@server.tool()
def doc_find_replace(
    old_text: str, new_text: str, match_case: bool = False
) -> dict[str, Any]:
    """Replaces every occurrence of the text in the document."""
    return call_bridge(
        "word",
        "find_replace",
        {"old_text": old_text, "new_text": new_text, "match_case": match_case},
    )


@server.tool()
def doc_add_bullet_list(items: list[Any]) -> dict[str, Any]:
    """Dodaje liste punktowana.

    'items' to teksty albo obiekty z poziomem: {"text": "Podpunkt", "level": 2}.
    """
    return call_bridge("word", "add_bullet_list", {"items": items})


@server.tool()
def doc_add_numbered_list(items: list[Any]) -> dict[str, Any]:
    """Adds a numbered list ('items' format as in doc_add_bullet_list)."""
    return call_bridge("word", "add_numbered_list", {"items": items})


@server.tool()
def doc_set_text_style(
    paragraph_index: int,
    font_name: str | None = None,
    font_size: float | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
) -> dict[str, Any]:
    """Formats the font of a whole paragraph (indexes from doc_get_outline)."""
    return call_bridge(
        "word",
        "set_text_style",
        {
            "paragraph_index": paragraph_index,
            "font_name": font_name,
            "font_size": font_size,
            "color": color,
            "bold": bold,
            "italic": italic,
            "underline": underline,
        },
    )


@server.tool()
def doc_set_paragraph_alignment(paragraph_index: int, alignment: str) -> dict[str, Any]:
    """Sets paragraph alignment: left, center, right or justify."""
    return call_bridge(
        "word",
        "set_paragraph_alignment",
        {"paragraph_index": paragraph_index, "alignment": alignment},
    )


@server.tool()
def doc_apply_style(paragraph_index: int, style_name: str) -> dict[str, Any]:
    """Nadaje akapitowi styl (Heading 1, Normal, Quote albo styl wlasny)."""
    return call_bridge(
        "word",
        "apply_style",
        {"paragraph_index": paragraph_index, "style_name": style_name},
    )


@server.tool()
def doc_set_page_margins(
    top: float, bottom: float, left: float, right: float, unit: str = "cm"
) -> dict[str, Any]:
    """Sets page margins. 'unit': cm, mm, in or pt."""
    return call_bridge(
        "word",
        "set_page_margins",
        {"top": top, "bottom": bottom, "left": left, "right": right, "unit": unit},
    )


@server.tool()
def doc_insert_image(
    image_path: str,
    width: float | None = None,
    height: float | None = None,
    position: str = "inline",
    unit: str = "pt",
    own_paragraph: bool = True,
) -> dict[str, Any]:
    """Inserts an image: 'inline' in the text or 'float' as a floating object."""
    return call_bridge(
        "word",
        "insert_image",
        {
            "image_path": image_path,
            "width": width,
            "height": height,
            "position": position,
            "unit": unit,
            "own_paragraph": own_paragraph,
        },
    )


@server.tool()
def doc_insert_table(
    rows: int,
    cols: int,
    data: list[list[Any]] | None = None,
    position: int | None = None,
) -> dict[str, Any]:
    """Inserts a bordered table; without 'position' it goes to the end of the document."""
    return call_bridge(
        "word",
        "insert_table",
        {"rows": rows, "cols": cols, "data": data, "position": position},
    )


@server.tool()
def doc_insert_header(text: str, section: int = 1) -> dict[str, Any]:
    """Sets the page header text."""
    return call_bridge("word", "insert_header", {"text": text, "section": section})


@server.tool()
def doc_insert_footer(text: str, section: int = 1) -> dict[str, Any]:
    """Sets the footer text."""
    return call_bridge("word", "insert_footer", {"text": text, "section": section})


@server.tool()
def doc_add_page_numbers(
    alignment: str = "center", first_page: bool = True
) -> dict[str, Any]:
    """Inserts page numbers into the footer (left, center or right)."""
    return call_bridge(
        "word",
        "add_page_numbers",
        {"alignment": alignment, "first_page": first_page},
    )


@server.tool()
def doc_insert_table_of_contents(levels: int = 3, position: Any = "start") -> dict[str, Any]:
    """Inserts a table of contents built from heading styles.

    'position': "start", "end" or a paragraph number to place it after -
    ten ostatni pozwala umiescic spis za strona tytulowa pracy dyplomowej.
    """
    return call_bridge(
        "word",
        "insert_table_of_contents",
        {"levels": levels, "position": position},
    )


def main() -> int:
    """MCP server entry point (stdio transport)."""
    logging.basicConfig(
        level=os.environ.get("OFFICE_MCP_LOG_LEVEL", "INFO"),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        stream=sys.stderr,
    )

    if sys.platform != "win32":
        print(
            "office-mcp runs on Windows only - Office automation is built on "
            f"COM, which does not exist on platform '{sys.platform}'.",
            file=sys.stderr,
        )
        return 1

    try:
        server.run()
    finally:
        client.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
