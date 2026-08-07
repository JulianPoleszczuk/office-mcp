"""Serwer MCP dla Microsoft Office 2019 (PowerPoint, Excel, Word).

Warstwa MCP jest cienka: kazde narzedzie zamienia argumenty na zadanie
protokolu Bridge i zwraca ustrukturyzowany JSON::

    {"ok": true,  "result": {...}}
    {"ok": false, "error": {"type": "ComConnectionError", "message": "..."}}

Zadne narzedzie nie propaguje wyjatku Pythona - awaria Office, brak Bridge
czy zla nazwa arkusza zawsze wracaja jako opisany blad.

Uruchomienie (stdio, tak jak startuje to Claude Desktop)::

    python server.py

Bridge (proces trzymajacy polaczenia COM) startuje automatycznie przy
pierwszym wywolaniu narzedzia. Mozna go tez uruchomic recznie:

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
    """Bridge nie odpowiada - nie udalo sie go uruchomic ani polaczyc."""


class BridgeClient:
    """Klient TCP Bridge: jedno polaczenie, autostart procesu, reconnect."""

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
        """Wysyla zadanie do Bridge; przy zerwanym polaczeniu ponawia raz."""
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
                        raise BridgeUnavailable("Bridge zamknal polaczenie")
                    return Response.decode(line)
                except (OSError, BridgeUnavailable) as exc:
                    last_error = exc
                    self._disconnect()
                    if attempt == 2:
                        break

        raise BridgeUnavailable(
            f"Brak polaczenia z Bridge na {self.host}:{self.port} ({last_error})"
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
            raise BridgeUnavailable("Nie udalo sie nawiazac polaczenia z Bridge")
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
        """Startuje Bridge w tle - Claude Desktop uruchamia tylko serwer MCP."""
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
            f"Bridge nie wystartowal w ciagu {BRIDGE_START_TIMEOUT:.0f}s - "
            f"uruchom go recznie: python -m bridge.main --port {self.port}"
        )


client = BridgeClient()
server = _McpServer(
    "office-mcp",
    instructions=(
        "Sterowanie otwartymi aplikacjami Microsoft Office 2019 na Windows przez COM. "
        "Narzedzia ppt_* obsluguja PowerPoint, xl_* Excel, doc_* Word. "
        "Aplikacje uruchamiaja sie leniwie, a zmiany widac na zywo w oknie Office. "
        "Wspolrzedne i rozmiary w PowerPoincie podaje sie w punktach (1 cm = 28.35 pt), "
        "indeksy slajdow i akapitow licza sie od 1."
    ),
)


def call_bridge(
    app: str,
    action: str,
    params: dict[str, Any] | None = None,
    keep_none: tuple[str, ...] = (),
) -> dict[str, Any]:
    """Wysyla akcje do Bridge i zwraca odpowiedz w stalym formacie JSON."""
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
    except Exception as exc:  # noqa: BLE001 - narzedzie MCP nigdy nie rzuca wyjatkiem
        logger.exception("Blad wywolania %s.%s", app, action)
        return {"ok": False, "error": {"type": type(exc).__name__, "message": str(exc)}}

    if response.ok:
        return {"ok": True, "result": response.result}
    return {"ok": False, "error": response.error}


@server.tool()
def office_status() -> dict[str, Any]:
    """Stan mostu COM: czy Bridge dziala i ktore aplikacje Office sa podlaczone."""
    apps = {}
    for app in ("powerpoint", "excel", "word"):
        apps[app] = call_bridge(app, "status")
    return {"ok": True, "result": {"bridge": client.status(), "apps": apps}}


@server.tool()
def ppt_create_presentation(path: str, template: str | None = None) -> dict[str, Any]:
    """Tworzy nowa prezentacje i zapisuje ja pod podana sciezka (.pptx).

    Opcjonalny 'template' to sciezka do pliku .potx/.thmx z motywem.
    """
    return call_bridge(
        "powerpoint", "create_presentation", {"path": path, "template": template}
    )


@server.tool()
def ppt_open_presentation(path: str) -> dict[str, Any]:
    """Otwiera istniejaca prezentacje; jesli jest juz otwarta, aktywuje jej okno."""
    return call_bridge("powerpoint", "open_presentation", {"path": path})


@server.tool()
def ppt_save(path: str | None = None) -> dict[str, Any]:
    """Zapisuje aktywna prezentacje; z 'path' robi zapis jako nowy plik."""
    return call_bridge("powerpoint", "save", {"path": path})


@server.tool()
def ppt_close(save: bool = True) -> dict[str, Any]:
    """Zamyka aktywna prezentacje, domyslnie zapisujac zmiany."""
    return call_bridge("powerpoint", "close", {"save": save})


@server.tool()
def ppt_get_presentation_info() -> dict[str, Any]:
    """Zwraca liczbe slajdow, rozmiar slajdu, nazwe motywu i sciezke pliku."""
    return call_bridge("powerpoint", "get_presentation_info")


@server.tool()
def ppt_get_slide_content(slide_index: int) -> dict[str, Any]:
    """Zwraca pelna zawartosc slajdu: ksztalty, ich pozycje, teksty i notatki."""
    return call_bridge("powerpoint", "get_slide_content", {"slide_index": slide_index})


@server.tool()
def ppt_list_slides() -> dict[str, Any]:
    """Lista slajdow z tytulami, nazwami ukladow i liczba ksztaltow."""
    return call_bridge("powerpoint", "list_slides")


@server.tool()
def ppt_add_slide(
    layout: str = "title_content",
    index: int | None = None,
    title: str | None = None,
) -> dict[str, Any]:
    """Dodaje slajd o wybranym ukladzie.

    Uklady: title, title_content, two_content, title_only, blank,
    section_header, comparison, picture_with_caption, chart, table.
    Bez 'index' slajd trafia na koniec prezentacji.
    """
    return call_bridge(
        "powerpoint", "add_slide", {"layout": layout, "index": index, "title": title}
    )


@server.tool()
def ppt_delete_slide(slide_index: int) -> dict[str, Any]:
    """Usuwa slajd o podanym numerze (liczac od 1)."""
    return call_bridge("powerpoint", "delete_slide", {"slide_index": slide_index})


@server.tool()
def ppt_duplicate_slide(slide_index: int) -> dict[str, Any]:
    """Duplikuje slajd - kopia trafia bezposrednio za oryginalem."""
    return call_bridge("powerpoint", "duplicate_slide", {"slide_index": slide_index})


@server.tool()
def ppt_reorder_slide(from_index: int, to_index: int) -> dict[str, Any]:
    """Przenosi slajd na inna pozycje w prezentacji."""
    return call_bridge(
        "powerpoint", "reorder_slide", {"from_index": from_index, "to_index": to_index}
    )


@server.tool()
def ppt_set_title(slide_index: int, text: str) -> dict[str, Any]:
    """Ustawia tytul slajdu; gdy uklad nie ma tytulu, wstawia pole tekstowe."""
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
    """Wstawia pole tekstowe. Wspolrzedne w punktach, slajd 16:9 ma 960x540 pt.

    'color' przyjmuje '#RRGGBB' albo nazwe koloru, 'align' to left/center/right/justify.
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
    """Wypelnia placeholder lista punktowana.

    'items' to teksty albo obiekty z poziomem wciecia:
    ["Punkt glowny", {"text": "Podpunkt", "level": 2}].
    'placeholder' to "content", "title" albo id ksztaltu.
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
    """Podmienia tekst w prezentacji (takze w tabelach i grupach ksztaltow).

    Bez 'slide_index' przeszukuje wszystkie slajdy.
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
    """Ustawia notatki prelegenta dla wskazanego slajdu."""
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
    """Formatuje tekst ksztaltu. 'shape_id' pobierz z ppt_get_slide_content."""
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
    """Nadaje prezentacji motyw z pliku .thmx/.potx albo z galerii motywow Office."""
    return call_bridge(
        "powerpoint", "apply_theme", {"theme_name_or_path": theme_name_or_path}
    )


@server.tool()
def ppt_set_background(
    slide_index: int,
    color: str | None = None,
    image_path: str | None = None,
) -> dict[str, Any]:
    """Ustawia tlo slajdu - jednolity kolor ('#RRGGBB') albo obraz z pliku."""
    return call_bridge(
        "powerpoint",
        "set_background",
        {"slide_index": slide_index, "color": color, "image_path": image_path},
    )


@server.tool()
def ppt_set_slide_layout(slide_index: int, layout_name: str) -> dict[str, Any]:
    """Zmienia uklad slajdu - po nazwie ukladu z wzorca albo nazwie standardowej."""
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
    """Wstawia obraz na slajd. Bez width/height zachowuje oryginalne proporcje."""
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
    """Wstawia wykres z danymi.

    'chart_type': bar, column, line, pie, area, scatter, doughnut, radar.
    'series_data': {"Wyniki 2024": [10, 20, 30]} albo lista serii
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
    """Wstawia tabele i wypelnia ja danymi (pierwszy wiersz zostaje pogrubiony)."""
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
    """Wstawia ksztalt: rectangle, rounded_rectangle, oval, triangle, diamond,
    star, arrow_right, callout, cloud, hexagon. 'fill_color'/'line_color'
    przyjmuja "none", zeby wylaczyc wypelnienie albo obrys."""
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
def ppt_delete_shape(slide_index: int, shape_id: Any) -> dict[str, Any]:
    """Usuwa ksztalt ze slajdu; 'shape_id' to id, nazwa ksztaltu albo skrot
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
    """Przesuwa, skaluje i obraca istniejacy ksztalt. Wspolrzedne w punktach
    (slajd 16:9 = 960 x 540 pt), 'rotation' w stopniach. Podaje sie tylko te
    pola, ktore maja sie zmienic."""
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
    """Zmienia warstwe ksztaltu: front (na wierzch), back (na spod),
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
    """Zapisuje slajd jako obraz - format wynika z rozszerzenia pliku
    (.png, .jpg, .gif, .bmp, .wmf, .emf). Bez podanych wymiarow obraz ma
    1920 px szerokosci. Sluzy do obejrzenia efektu i poprawienia ukladu."""
    return call_bridge(
        "powerpoint",
        "export_slide",
        {"slide_index": slide_index, "path": path, "width": width, "height": height},
    )


@server.tool()
def ppt_export_pdf(path: str, embed_fonts: bool = True) -> dict[str, Any]:
    """Eksportuje cala prezentacje do PDF-u; nie zmienia pliku otwartego
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
    """Animuje ksztalt na slajdzie. 'shape_id' to id, nazwa ksztaltu albo skrot
    'title'/'content'. Efekty: fade, fly, wipe, zoom, float, grow_and_turn,
    rise_up, split, wheel, bounce, spin, grow_shrink, teeter i inne.
    Wyzwalacze: on_click, with_previous, after_previous. 'level' = shape albo
    by_paragraph (tekst akapit po akapicie). 'duration' i 'delay' w sekundach."""
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
    """Zwraca animacje slajdu w kolejnosci odtwarzania wraz z przejsciem slajdu."""
    return call_bridge("powerpoint", "list_animations", {"slide_index": slide_index})


@server.tool()
def ppt_set_transition(
    effect: str = "fade",
    slide_index: int | None = None,
    duration: float | None = None,
    advance_on_click: bool = True,
    advance_after: float | None = None,
) -> dict[str, Any]:
    """Ustawia przejscie miedzy slajdami; bez 'slide_index' obejmuje cala
    prezentacje. Efekty: fade, fade_smoothly, push_left, wipe_right, cover_up,
    split_vertical_out, zoom_in, morph, honeycomb, gallery_left, cube_left,
    doors_vertical, curtains, prestige i inne. 'duration' w sekundach,
    'advance_after' wlacza automatyczna zmiane slajdu po zadanym czasie."""
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
    """Tworzy nowy skoroszyt i zapisuje go pod podana sciezka (.xlsx)."""
    return call_bridge("excel", "create_workbook", {"path": path})


@server.tool()
def xl_open_workbook(path: str) -> dict[str, Any]:
    """Otwiera istniejacy skoroszyt; jesli jest juz otwarty, aktywuje jego okno."""
    return call_bridge("excel", "open_workbook", {"path": path})


@server.tool()
def xl_save(path: str | None = None) -> dict[str, Any]:
    """Zapisuje aktywny skoroszyt; z 'path' robi zapis jako nowy plik."""
    return call_bridge("excel", "save", {"path": path})


@server.tool()
def xl_close(save: bool = True) -> dict[str, Any]:
    """Zamyka aktywny skoroszyt, domyslnie zapisujac zmiany."""
    return call_bridge("excel", "close", {"save": save})


@server.tool()
def xl_add_sheet(name: str, index: int | None = None) -> dict[str, Any]:
    """Dodaje arkusz o podanej nazwie; bez 'index' trafia on na koniec."""
    return call_bridge("excel", "add_sheet", {"name": name, "index": index})


@server.tool()
def xl_delete_sheet(name: str) -> dict[str, Any]:
    """Usuwa arkusz o podanej nazwie."""
    return call_bridge("excel", "delete_sheet", {"name": name})


@server.tool()
def xl_rename_sheet(old_name: str, new_name: str) -> dict[str, Any]:
    """Zmienia nazwe arkusza."""
    return call_bridge(
        "excel", "rename_sheet", {"old_name": old_name, "new_name": new_name}
    )


@server.tool()
def xl_get_workbook_info() -> dict[str, Any]:
    """Zwraca liste arkuszy z ich zakresami danych, aktywny arkusz i sciezke pliku."""
    return call_bridge("excel", "get_workbook_info")


@server.tool()
def xl_get_range_values(sheet: str, range_ref: str) -> dict[str, Any]:
    """Odczytuje wartosci zakresu (np. "A1:D10") jako tablice 2D."""
    return call_bridge("excel", "get_range_values", {"sheet": sheet, "range_ref": range_ref})


@server.tool()
def xl_get_used_range(sheet: str) -> dict[str, Any]:
    """Zwraca faktycznie wypelniony obszar arkusza razem z danymi."""
    return call_bridge("excel", "get_used_range", {"sheet": sheet})


@server.tool()
def xl_set_cell(sheet: str, cell_ref: str, value: Any) -> dict[str, Any]:
    """Wpisuje wartosc do komorki (np. sheet="Budzet", cell_ref="B2")."""
    return call_bridge(
        "excel",
        "set_cell",
        {"sheet": sheet, "cell_ref": cell_ref, "value": value},
        keep_none=("value",),
    )


@server.tool()
def xl_set_range(sheet: str, start_cell: str, values_2d: list[list[Any]]) -> dict[str, Any]:
    """Wkleja macierz danych naraz, zaczynajac od 'start_cell'.

    Duzo szybsze niz wpisywanie komorka po komorce.
    """
    return call_bridge(
        "excel",
        "set_range",
        {"sheet": sheet, "start_cell": start_cell, "values_2d": values_2d},
    )


@server.tool()
def xl_set_formula(sheet: str, cell_ref: str, formula: str) -> dict[str, Any]:
    """Wpisuje formule (np. "=SUM(A1:A10)") i zwraca wyliczony wynik."""
    return call_bridge(
        "excel", "set_formula", {"sheet": sheet, "cell_ref": cell_ref, "formula": formula}
    )


@server.tool()
def xl_clear_range(sheet: str, range_ref: str, contents_only: bool = True) -> dict[str, Any]:
    """Czysci zakres - domyslnie same wartosci, opcjonalnie takze formatowanie."""
    return call_bridge(
        "excel",
        "clear_range",
        {"sheet": sheet, "range_ref": range_ref, "contents_only": contents_only},
    )


@server.tool()
def xl_insert_rows(sheet: str, start_row: int, count: int = 1) -> dict[str, Any]:
    """Wstawia wiersze, przesuwajac istniejace w dol."""
    return call_bridge(
        "excel", "insert_rows", {"sheet": sheet, "start_row": start_row, "count": count}
    )


@server.tool()
def xl_delete_rows(sheet: str, start_row: int, count: int = 1) -> dict[str, Any]:
    """Usuwa wiersze, przesuwajac pozostale w gore."""
    return call_bridge(
        "excel", "delete_rows", {"sheet": sheet, "start_row": start_row, "count": count}
    )


@server.tool()
def xl_insert_columns(sheet: str, start_col: str | int, count: int = 1) -> dict[str, Any]:
    """Wstawia kolumny; 'start_col' przyjmuje litere ("C") albo numer (3)."""
    return call_bridge(
        "excel",
        "insert_columns",
        {"sheet": sheet, "start_col": start_col, "count": count},
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
    """Formatuje zakres komorek.

    'number_format' to maska Excela, np. "0.00", "# ##0 zl", "0%".
    Kolory przyjmuja '#RRGGBB' albo nazwe (red, blue, green...).
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
    """Ustawia szerokosc kolumny; width="auto" dopasowuje ja do zawartosci."""
    return call_bridge(
        "excel", "set_column_width", {"sheet": sheet, "column": column, "width": width}
    )


@server.tool()
def xl_merge_cells(sheet: str, range_ref: str, center: bool = True) -> dict[str, Any]:
    """Scala komorki zakresu, domyslnie centrujac zawartosc."""
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
    """Zamraza wiersze i kolumny powyzej/na lewo od komorki (np. "A2")."""
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
    """Wstawia wykres oparty o zakres danych z tego samego arkusza.

    'chart_type': column, bar, line, pie, area, scatter, doughnut, radar.
    Pozycja i rozmiar w punktach.
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
    """Zamienia zakres w natywna tabele Excela (z filtrami i stylem)."""
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
    """Tworzy tabele przestawna z zakresu zrodlowego (pierwszy wiersz to naglowki).

    'values' to nazwy pol albo obiekty {"field": "Kwota", "function": "average"}.
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
def doc_create_document(path: str, template: str | None = None) -> dict[str, Any]:
    """Tworzy nowy dokument i zapisuje go pod podana sciezka (.docx).

    Opcjonalny 'template' to sciezka do pliku .dotx.
    """
    return call_bridge("word", "create_document", {"path": path, "template": template})


@server.tool()
def doc_open_document(path: str) -> dict[str, Any]:
    """Otwiera istniejacy dokument; jesli jest juz otwarty, aktywuje jego okno."""
    return call_bridge("word", "open_document", {"path": path})


@server.tool()
def doc_save(path: str | None = None) -> dict[str, Any]:
    """Zapisuje aktywny dokument; z 'path' robi zapis jako nowy plik."""
    return call_bridge("word", "save", {"path": path})


@server.tool()
def doc_close(save: bool = True) -> dict[str, Any]:
    """Zamyka aktywny dokument, domyslnie zapisujac zmiany."""
    return call_bridge("word", "close", {"save": save})


@server.tool()
def doc_get_document_info() -> dict[str, Any]:
    """Zwraca liczbe stron, slow i znakow, nazwe szablonu oraz sciezke pliku."""
    return call_bridge("word", "get_document_info")


@server.tool()
def doc_get_full_text() -> dict[str, Any]:
    """Zwraca caly tekst dokumentu."""
    return call_bridge("word", "get_full_text")


@server.tool()
def doc_get_outline() -> dict[str, Any]:
    """Zwraca strukture naglowkow (poziom, tekst, indeks akapitu)."""
    return call_bridge("word", "get_outline")


@server.tool()
def doc_add_paragraph(text: str, style: str | None = None) -> dict[str, Any]:
    """Dopisuje akapit na koncu dokumentu.

    'style' przyjmuje nazwy angielskie (Normal, Heading 1, Quote, Caption)
    takze w polskiej wersji Worda.
    """
    return call_bridge("word", "add_paragraph", {"text": text, "style": style})


@server.tool()
def doc_add_heading(text: str, level: int = 1) -> dict[str, Any]:
    """Dopisuje naglowek poziomu 1-9 (styl Heading N)."""
    return call_bridge("word", "add_heading", {"text": text, "level": level})


@server.tool()
def doc_insert_page_break() -> dict[str, Any]:
    """Wstawia twardy podzial strony na koncu dokumentu."""
    return call_bridge("word", "insert_page_break")


@server.tool()
def doc_find_replace(
    old_text: str, new_text: str, match_case: bool = False
) -> dict[str, Any]:
    """Podmienia wszystkie wystapienia tekstu w dokumencie."""
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
    """Dodaje liste numerowana (format 'items' jak w doc_add_bullet_list)."""
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
    """Formatuje czcionke calego akapitu (indeksy z doc_get_outline)."""
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
    """Ustawia wyrownanie akapitu: left, center, right albo justify."""
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
    """Ustawia marginesy strony. 'unit': cm, mm, in albo pt."""
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
) -> dict[str, Any]:
    """Wstawia obraz: 'inline' w tekscie albo 'float' jako obiekt plywajacy."""
    return call_bridge(
        "word",
        "insert_image",
        {
            "image_path": image_path,
            "width": width,
            "height": height,
            "position": position,
        },
    )


@server.tool()
def doc_insert_table(
    rows: int,
    cols: int,
    data: list[list[Any]] | None = None,
    position: int | None = None,
) -> dict[str, Any]:
    """Wstawia tabele z obramowaniem; bez 'position' trafia na koniec dokumentu."""
    return call_bridge(
        "word",
        "insert_table",
        {"rows": rows, "cols": cols, "data": data, "position": position},
    )


@server.tool()
def doc_insert_header(text: str, section: int = 1) -> dict[str, Any]:
    """Ustawia tekst naglowka strony."""
    return call_bridge("word", "insert_header", {"text": text, "section": section})


@server.tool()
def doc_insert_footer(text: str, section: int = 1) -> dict[str, Any]:
    """Ustawia tekst stopki."""
    return call_bridge("word", "insert_footer", {"text": text, "section": section})


@server.tool()
def doc_add_page_numbers(
    alignment: str = "center", first_page: bool = True
) -> dict[str, Any]:
    """Wstawia numery stron w stopce (left, center albo right)."""
    return call_bridge(
        "word",
        "add_page_numbers",
        {"alignment": alignment, "first_page": first_page},
    )


@server.tool()
def doc_insert_table_of_contents(levels: int = 3, position: str = "start") -> dict[str, Any]:
    """Wstawia spis tresci zbudowany ze stylow naglowkow.

    'position': "start" (poczatek dokumentu) albo "end".
    """
    return call_bridge(
        "word",
        "insert_table_of_contents",
        {"levels": levels, "position": position},
    )


def main() -> int:
    """Punkt wejscia serwera MCP (transport stdio)."""
    logging.basicConfig(
        level=os.environ.get("OFFICE_MCP_LOG_LEVEL", "INFO"),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        stream=sys.stderr,
    )

    if sys.platform != "win32":
        print(
            "office-mcp dziala tylko na Windows - automatyzacja Office opiera sie na "
            f"COM, ktorego nie ma na platformie '{sys.platform}'.",
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
