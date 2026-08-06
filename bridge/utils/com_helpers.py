"""Konwersje miedzy swiatem Pythona a COM-em Office.

Zebrane tu sa rzeczy, ktore inaczej powtarzalyby sie w kazdym kontrolerze:
kolory (Office trzyma je jako BGR, nie RGB), jednostki (punkty / centymetry /
cale / EMU), stale numeryczne enumow Office oraz oczyszczanie wartosci
zwracanych przez COM do postaci serializowalnej w JSON.
"""

from __future__ import annotations

import datetime as _dt
import os
from typing import Any, Iterable, Sequence

try:
    import pywintypes

    com_error = pywintypes.com_error
    COM_AVAILABLE = True
except ImportError:  # pragma: no cover - platformy bez pywin32 (CI, testy)

    class com_error(Exception):  # type: ignore[no-redef]
        """Zaslepka uzywana, gdy pywin32 nie jest dostepny (testy poza Windows)."""

    COM_AVAILABLE = False


POINTS_PER_INCH = 72.0
POINTS_PER_CM = 28.3464567
EMU_PER_POINT = 12700

MSO_TRUE = -1
MSO_FALSE = 0

NAMED_COLORS: dict[str, tuple[int, int, int]] = {
    "black": (0, 0, 0),
    "white": (255, 255, 255),
    "red": (255, 0, 0),
    "green": (0, 176, 80),
    "blue": (0, 112, 192),
    "yellow": (255, 255, 0),
    "orange": (255, 153, 0),
    "purple": (112, 48, 160),
    "pink": (255, 102, 204),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "lightgray": (217, 217, 217),
    "lightgrey": (217, 217, 217),
    "darkgray": (64, 64, 64),
    "darkgrey": (64, 64, 64),
    "brown": (132, 60, 12),
    "navy": (0, 32, 96),
    "teal": (0, 128, 128),
    "gold": (255, 192, 0),
    "silver": (191, 191, 191),
    "czarny": (0, 0, 0),
    "bialy": (255, 255, 255),
    "czerwony": (255, 0, 0),
    "zielony": (0, 176, 80),
    "niebieski": (0, 112, 192),
    "zolty": (255, 255, 0),
    "pomaranczowy": (255, 153, 0),
    "szary": (128, 128, 128),
}

PP_LAYOUTS: dict[str, int] = {
    "title": 1,
    "title_content": 2,
    "text": 2,
    "two_content": 3,
    "two_column_text": 3,
    "table": 4,
    "text_and_chart": 5,
    "chart_and_text": 6,
    "org_chart": 7,
    "chart": 8,
    "title_only": 11,
    "blank": 12,
    "text_and_object": 13,
    "large_object": 15,
    "object": 16,
    "four_objects": 24,
    "vertical_text": 25,
    "vertical_title_and_text": 27,
    "two_objects": 29,
    "custom": 32,
    "section_header": 33,
    "comparison": 34,
    "content_with_caption": 35,
    "picture_with_caption": 36,
}

CHART_TYPES: dict[str, int] = {
    "column": 51,
    "column_clustered": 51,
    "bar": 57,
    "bar_clustered": 57,
    "column_stacked": 52,
    "bar_stacked": 58,
    "line": 4,
    "line_markers": 65,
    "pie": 5,
    "pie_3d": -4102,
    "doughnut": -4120,
    "area": 1,
    "area_stacked": 76,
    "scatter": -4169,
    "radar": -4151,
    "bubble": 15,
    "column_3d": -4100,
    "stock": 88,
}

SHAPE_TYPES: dict[str, int] = {
    "rectangle": 1,
    "rounded_rectangle": 5,
    "oval": 9,
    "circle": 9,
    "triangle": 7,
    "right_triangle": 8,
    "diamond": 4,
    "pentagon": 51,
    "hexagon": 10,
    "chevron": 52,
    "star": 92,
    "arrow_right": 33,
    "arrow_left": 34,
    "arrow_up": 35,
    "arrow_down": 36,
    "callout": 105,
    "cloud": 179,
    "heart": 21,
    "smiley": 17,
    "plus": 11,
    "line_shape": 20,
}

WD_ALIGNMENTS: dict[str, int] = {
    "left": 0,
    "center": 1,
    "centre": 1,
    "right": 2,
    "justify": 3,
}

WD_BUILTIN_STYLES: dict[str, int] = {
    "normal": -1,
    "heading 1": -2,
    "heading 2": -3,
    "heading 3": -4,
    "heading 4": -5,
    "heading 5": -6,
    "heading 6": -7,
    "heading 7": -8,
    "heading 8": -9,
    "heading 9": -10,
    "title": -63,
    "subtitle": -74,
    "list bullet": -48,
    "list number": -49,
    "caption": -35,
    "quote": -88,
}

XL_COMPARISON_OPERATORS: dict[str, int] = {
    "between": 1,
    "not_between": 2,
    "equal": 3,
    "not_equal": 4,
    "greater": 5,
    "less": 6,
    "greater_equal": 7,
    "less_equal": 8,
}

XL_SAVE_FORMATS: dict[str, int] = {
    ".xlsx": 51,
    ".xlsm": 52,
    ".xlsb": 50,
    ".xls": 56,
    ".csv": 6,
    ".pdf": 57,
}

PP_SAVE_FORMATS: dict[str, int] = {
    ".pptx": 24,
    ".pptm": 25,
    ".ppt": 1,
    ".pdf": 32,
    ".potx": 27,
}

WD_SAVE_FORMATS: dict[str, int] = {
    ".docx": 16,
    ".docm": 13,
    ".doc": 0,
    ".pdf": 17,
    ".txt": 2,
    ".rtf": 6,
}


def parse_color(value: Any) -> int:
    """Zamienia kolor podany po ludzku na liczbe BGR oczekiwana przez Office.

    Akceptuje ``"#RRGGBB"``, ``"RRGGBB"``, nazwe (``"red"``, ``"czerwony"``),
    krotke/liste ``(r, g, b)`` oraz gotowa liczbe calkowita RGB.
    """
    if value is None:
        raise ValueError("Kolor nie moze byc pusty")

    if isinstance(value, (tuple, list)):
        if len(value) != 3:
            raise ValueError("Kolor jako krotka musi miec dokladnie 3 skladowe RGB")
        r, g, b = (int(component) for component in value)
    elif isinstance(value, int):
        r, g, b = (value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF
    elif isinstance(value, str):
        text = value.strip().lower()
        if text in NAMED_COLORS:
            r, g, b = NAMED_COLORS[text]
        else:
            hex_text = text.lstrip("#")
            if len(hex_text) == 3:
                hex_text = "".join(ch * 2 for ch in hex_text)
            if len(hex_text) != 6:
                raise ValueError(f"Nieznany kolor: {value!r}")
            try:
                r, g, b = (int(hex_text[i : i + 2], 16) for i in (0, 2, 4))
            except ValueError as exc:
                raise ValueError(f"Nieznany kolor: {value!r}") from exc
    else:
        raise ValueError(f"Nieobslugiwany typ koloru: {type(value).__name__}")

    for component in (r, g, b):
        if not 0 <= component <= 255:
            raise ValueError("Skladowe koloru musza miescic sie w zakresie 0-255")

    return (b << 16) | (g << 8) | r


def bgr_to_hex(value: Any) -> str | None:
    """Odwrotnosc :func:`parse_color` - z liczby BGR robi ``#RRGGBB``."""
    if value is None:
        return None
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    if number < 0:
        return None
    b, g, r = (number >> 16) & 0xFF, (number >> 8) & 0xFF, number & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def points(value: Any, unit: str = "pt") -> float:
    """Przelicza wartosc na punkty (jednostka pozycjonowania w Office)."""
    number = float(value)
    unit = (unit or "pt").lower()
    if unit in ("pt", "point", "points"):
        return number
    if unit in ("cm", "centimeter", "centimeters"):
        return number * POINTS_PER_CM
    if unit in ("mm", "millimeter", "millimeters"):
        return number * POINTS_PER_CM / 10
    if unit in ("in", "inch", "inches"):
        return number * POINTS_PER_INCH
    if unit in ("emu",):
        return number / EMU_PER_POINT
    raise ValueError(f"Nieznana jednostka: {unit}")


def points_to_emu(value: float) -> int:
    return int(round(float(value) * EMU_PER_POINT))


def emu_to_points(value: float) -> float:
    return float(value) / EMU_PER_POINT


def to_python(value: Any) -> Any:
    """Sprowadza wartosc z COM do typu, ktory da sie zserializowac do JSON."""
    if value is None or isinstance(value, (bool, int, str)):
        return value

    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return None
        return round(value, 6) if value % 1 else value

    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return value.isoformat()

    if isinstance(value, (tuple, list)):
        return [to_python(item) for item in value]

    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:  # noqa: BLE001
            return str(value)

    return str(value)


def to_matrix(values: Any) -> list[list[Any]]:
    """Normalizuje dane wejsciowe do prostokatnej macierzy 2D (lista list)."""
    if values is None:
        return []

    if not isinstance(values, (list, tuple)):
        return [[values]]

    rows: list[list[Any]] = []
    for row in values:
        if isinstance(row, (list, tuple)):
            rows.append(list(row))
        else:
            rows.append([row])

    width = max((len(row) for row in rows), default=0)
    for row in rows:
        row.extend([None] * (width - len(row)))
    return rows


def from_com_matrix(value: Any) -> list[list[Any]]:
    """Zamienia wynik ``Range.Value`` (skalar / krotka krotek) na liste list."""
    if value is None:
        return [[None]]
    if not isinstance(value, (tuple, list)):
        return [[to_python(value)]]
    if value and not isinstance(value[0], (tuple, list)):
        return [[to_python(item) for item in value]]
    return [[to_python(cell) for cell in row] for row in value]


def to_com_matrix(values: Sequence[Sequence[Any]]) -> tuple[tuple[Any, ...], ...]:
    """Zamienia macierz Pythona na krotke krotek - format przyjmowany przez Excela."""
    return tuple(tuple(row) for row in to_matrix(values))


def normalize_path(path: str, must_exist: bool = False) -> str:
    """Rozwija ``~``, zmienne srodowiskowe i zwraca sciezke absolutna Windows."""
    if not path or not isinstance(path, str):
        raise ValueError("Sciezka musi byc niepustym tekstem")
    expanded = os.path.abspath(os.path.expandvars(os.path.expanduser(path.strip())))
    if must_exist and not os.path.isfile(expanded):
        raise FileNotFoundError(expanded)
    return expanded


def save_format_for(path: str, formats: dict[str, int], default: int) -> int:
    """Dobiera stala formatu zapisu Office na podstawie rozszerzenia pliku."""
    extension = os.path.splitext(path)[1].lower()
    return formats.get(extension, default)


def lookup_constant(
    name: Any,
    mapping: dict[str, int],
    label: str,
) -> int:
    """Tlumaczy przyjazna nazwe (``"bar"``, ``"blank"``) na stala numeryczna Office."""
    if isinstance(name, bool):
        raise ValueError(f"Nieprawidlowa wartosc dla {label}: {name!r}")
    if isinstance(name, int):
        return name
    if not isinstance(name, str):
        raise ValueError(f"Nieprawidlowa wartosc dla {label}: {name!r}")

    key = name.strip().lower().replace("-", "_").replace(" ", "_")
    if key in mapping:
        return mapping[key]

    available = ", ".join(sorted(mapping))
    raise ValueError(f"Nieznane {label}: {name!r}. Dostepne: {available}")


def chunked(items: Iterable[Any], size: int) -> Iterable[list[Any]]:
    """Dzieli iterowalne na porcje o zadanym rozmiarze."""
    chunk: list[Any] = []
    for item in items:
        chunk.append(item)
        if len(chunk) >= size:
            yield chunk
            chunk = []
    if chunk:
        yield chunk


def column_letter(index: int) -> str:
    """Zamienia numer kolumny (1-based) na oznaczenie literowe Excela."""
    if index < 1:
        raise ValueError("Numer kolumny musi byc >= 1")
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def column_index(letter: str) -> int:
    """Zamienia oznaczenie literowe kolumny Excela na numer (1-based)."""
    letters = str(letter).strip().upper()
    if not letters.isalpha():
        raise ValueError(f"Nieprawidlowe oznaczenie kolumny: {letter!r}")
    result = 0
    for char in letters:
        result = result * 26 + (ord(char) - 64)
    return result
