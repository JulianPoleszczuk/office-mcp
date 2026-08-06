"""Kontroler Worda - tresc, style, naglowki i obiekty osadzone przez COM.

Akapity indeksuje sie od 1, tak jak w kolekcji ``Document.Paragraphs``.
Nazwy stylow mozna podawac po angielsku (``"Heading 1"``, ``"Normal"``) nawet
w polskiej wersji Worda - kontroler mapuje je na stale wbudowane.
"""

from __future__ import annotations

import os
from typing import Any

from bridge.controllers.base import BaseController, action, is_connection_error
from bridge.utils.com_helpers import (
    WD_ALIGNMENTS,
    WD_BUILTIN_STYLES,
    WD_SAVE_FORMATS,
    com_error,
    parse_color,
    points,
    save_format_for,
    to_python,
)
from bridge.utils.errors import (
    DocumentNotFoundError,
    InvalidReferenceError,
    UnsupportedOperationError,
)

WD_COLLAPSE_END = 0
WD_PAGE_BREAK = 7
WD_REPLACE_ALL = 2
WD_FIND_CONTINUE = 1
WD_HEADER_FOOTER_PRIMARY = 1
WD_STATISTIC_WORDS = 0
WD_STATISTIC_PAGES = 2
WD_STATISTIC_CHARACTERS = 3
WD_OUTLINE_BODY_TEXT = 10
WD_LINE_STYLE_SINGLE = 1


class WordController(BaseController):
    """Akcje ``doc_*`` - operacje na zywej instancji Worda."""

    APP_KEY = "word"
    DISPLAY_NAME = "Word"
    ALERTS_OFF = 0

    def document(self) -> Any:
        """Aktywny dokument albo czytelny blad, gdy nic nie jest otwarte."""
        app = self.app
        if app.Documents.Count == 0:
            raise DocumentNotFoundError(
                "Brak otwartego dokumentu - uzyj doc_create_document albo doc_open_document"
            )
        try:
            return app.ActiveDocument
        except com_error:
            return app.Documents(app.Documents.Count)

    def paragraph(self, paragraph_index: Any) -> Any:
        """Akapit o zadanym indeksie (1-based) z walidacja zakresu."""
        document = self.document()
        index = self.require_index(
            paragraph_index, document.Paragraphs.Count, "paragraph_index"
        )
        return document.Paragraphs(index)

    def _document_summary(self, document: Any) -> dict[str, Any]:
        return {
            "name": to_python(document.Name),
            "path": to_python(document.FullName) if document.Path else None,
            "paragraph_count": int(document.Paragraphs.Count),
            "saved": bool(document.Saved),
        }

    def _end_range(self, document: Any) -> Any:
        """Zakres ustawiony na sam koniec dokumentu."""
        target = document.Content
        target.Collapse(WD_COLLAPSE_END)
        return target

    def _append_paragraph(self, document: Any, text: str) -> Any:
        """Dopisuje akapit z tekstem na koncu dokumentu i zwraca go.

        Swiadomie nie uzywamy ``Paragraphs.Add`` ani przypisania do
        ``Range.Text``: pierwsze wstawia akapit w miejscu zaznaczenia, a drugie
        nadpisuje znak konca akapitu i skleja sasiednie akapity w jeden.
        Pusty akapit na koncu dokumentu jest wykorzystywany ponownie.
        """
        content = document.Content
        content.Collapse(WD_COLLAPSE_END)

        if self._has_text(document.Paragraphs(document.Paragraphs.Count)):
            content.InsertParagraphAfter()

        content.InsertAfter(str(text))
        return document.Paragraphs(document.Paragraphs.Count)

    @staticmethod
    def _has_text(paragraph: Any) -> bool:
        """Czy akapit zawiera cokolwiek poza znakami konca akapitu i komorki."""
        raw = str(paragraph.Range.Text)
        return bool(raw.replace("\r", "").replace("\x07", "").strip())

    def _apply_named_style(self, target: Any, style_name: str) -> str:
        """Ustawia styl po nazwie lokalnej albo po stalej wbudowanej Worda."""
        wanted = str(style_name).strip()

        try:
            target.Style = wanted
            return wanted
        except com_error as exc:
            if is_connection_error(exc):
                raise

        builtin = WD_BUILTIN_STYLES.get(wanted.lower())
        if builtin is None:
            raise InvalidReferenceError(
                f"Nieznany styl '{style_name}'. Uzyj nazwy z Worda albo jednej z: "
                f"{', '.join(sorted(WD_BUILTIN_STYLES))}"
            )

        try:
            target.Style = builtin
        except com_error as exc:
            raise InvalidReferenceError(
                f"Nie udalo sie zastosowac stylu '{style_name}'"
            ) from exc
        return wanted

    @action("create_document")
    def create_document(self, path: str, template: str | None = None) -> dict[str, Any]:
        """Tworzy dokument (opcjonalnie z szablonu .dotx) i zapisuje go."""
        target = self.resolve_target_path(path)

        if template:
            document = self.app.Documents.Add(Template=self.resolve_existing_path(template))
        else:
            document = self.app.Documents.Add()

        with self.alerts_suppressed():
            document.SaveAs2(
                target, save_format_for(target, WD_SAVE_FORMATS, WD_SAVE_FORMATS[".docx"])
            )
        return self._document_summary(document)

    @action("open_document")
    def open_document(self, path: str) -> dict[str, Any]:
        """Otwiera plik albo aktywuje go, jesli jest juz otwarty."""
        target = self.resolve_existing_path(path)
        app = self.app

        for index in range(1, app.Documents.Count + 1):
            document = app.Documents(index)
            if os.path.normcase(str(document.FullName)) == os.path.normcase(target):
                try:
                    document.Activate()
                except com_error:
                    pass
                return {**self._document_summary(document), "already_open": True}

        document = app.Documents.Open(target)
        return {**self._document_summary(document), "already_open": False}

    @action("save")
    def save(self, path: str | None = None) -> dict[str, Any]:
        """Zapisuje dokument albo zapisuje go jako nowy plik."""
        document = self.document()

        if path:
            target = self.resolve_target_path(path)
            with self.alerts_suppressed():
                document.SaveAs2(
                    target,
                    save_format_for(target, WD_SAVE_FORMATS, WD_SAVE_FORMATS[".docx"]),
                )
        elif not document.Path:
            raise InvalidReferenceError(
                "Dokument nie ma jeszcze pliku - podaj parametr path"
            )
        else:
            document.Save()

        return self._document_summary(document)

    @action("close")
    def close(self, save: bool = True) -> dict[str, Any]:
        """Zamyka dokument, opcjonalnie zapisujac zmiany."""
        document = self.document()
        name = str(document.Name)

        if save:
            if not document.Path:
                raise InvalidReferenceError(
                    "Dokument nie byl zapisany - najpierw doc_save z parametrem path"
                )
            document.Save()

        with self.alerts_suppressed():
            document.Close(SaveChanges=bool(save))

        return {"closed": name, "saved": bool(save)}

    @action("get_document_info")
    def get_document_info(self) -> dict[str, Any]:
        """Metadane dokumentu: liczba stron i slow, szablon, sciezka."""
        document = self.document()
        info = self._document_summary(document)

        for key, statistic in (
            ("pages", WD_STATISTIC_PAGES),
            ("words", WD_STATISTIC_WORDS),
            ("characters", WD_STATISTIC_CHARACTERS),
        ):
            try:
                info[key] = int(document.ComputeStatistics(statistic))
            except com_error:
                info[key] = None

        try:
            info["template"] = to_python(document.AttachedTemplate.Name)
        except com_error:
            info["template"] = None

        try:
            info["first_paragraph_style"] = to_python(document.Paragraphs(1).Style.NameLocal)
        except com_error:
            info["first_paragraph_style"] = None

        return info

    @action("get_full_text")
    def get_full_text(self) -> dict[str, Any]:
        """Zwraca caly tekst dokumentu (akapity rozdzielone znakiem nowej linii)."""
        document = self.document()
        text = str(document.Content.Text).replace("\r\x07", "\n").replace("\r", "\n")

        return {
            "text": text,
            "characters": len(text),
            "paragraph_count": int(document.Paragraphs.Count),
        }

    @action("get_outline")
    def get_outline(self) -> dict[str, Any]:
        """Buduje drzewo naglowkow na podstawie poziomow konspektu."""
        document = self.document()
        headings = []

        for index in range(1, document.Paragraphs.Count + 1):
            paragraph = document.Paragraphs(index)
            try:
                level = int(paragraph.OutlineLevel)
            except com_error:
                continue

            if level >= WD_OUTLINE_BODY_TEXT:
                continue

            text = str(paragraph.Range.Text).replace("\r", "").replace("\x07", "").strip()
            if not text:
                continue

            headings.append(
                {
                    "paragraph_index": index,
                    "level": level,
                    "text": text,
                    "style": to_python(paragraph.Style.NameLocal),
                }
            )

        return {"headings": headings, "count": len(headings)}

    @action("add_paragraph")
    def add_paragraph(self, text: str, style: str | None = None) -> dict[str, Any]:
        """Dopisuje akapit na koncu dokumentu, opcjonalnie z wybranym stylem."""
        document = self.document()
        paragraph = self._append_paragraph(document, text)

        applied_style = None
        if style:
            applied_style = self._apply_named_style(paragraph.Range, style)

        return {
            "paragraph_index": int(document.Paragraphs.Count),
            "style": applied_style,
            "characters": len(str(text)),
        }

    @action("add_heading")
    def add_heading(self, text: str, level: int = 1) -> dict[str, Any]:
        """Dopisuje naglowek poziomu 1-9."""
        try:
            heading_level = int(level)
        except (TypeError, ValueError) as exc:
            raise InvalidReferenceError("Poziom naglowka musi byc liczba 1-9") from exc

        if not 1 <= heading_level <= 9:
            raise InvalidReferenceError("Poziom naglowka musi miescic sie w zakresie 1-9")

        result = self.add_paragraph(text, style=f"Heading {heading_level}")
        result["level"] = heading_level
        return result

    @action("insert_page_break")
    def insert_page_break(self) -> dict[str, Any]:
        """Wstawia twardy podzial strony na koncu dokumentu."""
        document = self.document()
        self._end_range(document).InsertBreak(WD_PAGE_BREAK)
        return {"paragraph_count": int(document.Paragraphs.Count)}

    @action("find_replace")
    def find_replace(
        self, old_text: str, new_text: str, match_case: bool = False
    ) -> dict[str, Any]:
        """Podmienia tekst w calym dokumencie i zwraca liczbe trafien.

        Wszystkie parametry wyszukiwania ida w jednym wywolaniu ``Execute`` -
        ustawianie ich jako wlasciwosci obiektu ``Find`` przy poznym wiazaniu
        COM zwraca sukces, ale nie podmienia tekstu.

        Przy ``match_case=False`` Word dopasowuje wielkosc liter wstawianego
        tekstu do znalezionego (tak samo jak okno Znajdz i zamien).
        """
        if not old_text:
            raise InvalidReferenceError("Parametr old_text nie moze byc pusty")

        document = self.document()
        content = str(document.Content.Text)
        occurrences = (
            content.count(old_text)
            if match_case
            else content.lower().count(str(old_text).lower())
        )

        replaced = bool(
            document.Content.Find.Execute(
                FindText=str(old_text),
                MatchCase=bool(match_case),
                MatchWholeWord=False,
                MatchWildcards=False,
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Forward=True,
                Wrap=WD_FIND_CONTINUE,
                Format=False,
                ReplaceWith=str(new_text),
                Replace=WD_REPLACE_ALL,
            )
        )

        return {
            "replacements": occurrences if replaced else 0,
            "old_text": str(old_text),
            "new_text": str(new_text),
            "match_case": bool(match_case),
        }

    @action("add_bullet_list")
    def add_bullet_list(self, items: list[Any]) -> dict[str, Any]:
        """Dodaje liste punktowana (obsluguje poziomy zagniezdzenia)."""
        return self._add_list(items, numbered=False)

    @action("add_numbered_list")
    def add_numbered_list(self, items: list[Any]) -> dict[str, Any]:
        """Dodaje liste numerowana (obsluguje poziomy zagniezdzenia)."""
        return self._add_list(items, numbered=True)

    def _add_list(self, items: list[Any], numbered: bool) -> dict[str, Any]:
        if not isinstance(items, list) or not items:
            raise InvalidReferenceError("Lista 'items' nie moze byc pusta")

        document = self.document()
        entries = []
        for item in items:
            if isinstance(item, dict):
                text = str(item.get("text", ""))
                level = int(item.get("level", 1))
            elif isinstance(item, (list, tuple)) and len(item) == 2:
                text, level = str(item[0]), int(item[1])
            else:
                text, level = str(item), 1
            entries.append((text, max(1, min(level, 9))))

        first_index = None
        paragraphs = []
        for text, level in entries:
            paragraph = self._append_paragraph(document, text)
            paragraphs.append((paragraph, level))
            if first_index is None:
                first_index = int(document.Paragraphs.Count)

        start = paragraphs[0][0].Range.Start
        end = paragraphs[-1][0].Range.End
        list_range = document.Range(start, end)

        if numbered:
            list_range.ListFormat.ApplyNumberDefault()
        else:
            list_range.ListFormat.ApplyBulletDefault()

        for paragraph, level in paragraphs:
            if level > 1:
                try:
                    paragraph.Range.ListFormat.ListLevelNumber = level
                except com_error:
                    pass

        return {
            "items": len(entries),
            "numbered": bool(numbered),
            "first_paragraph_index": first_index,
            "last_paragraph_index": int(document.Paragraphs.Count),
        }

    @action("set_text_style")
    def set_text_style(
        self,
        paragraph_index: int,
        font_name: str | None = None,
        font_size: float | None = None,
        color: Any = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
    ) -> dict[str, Any]:
        """Formatuje czcionke calego akapitu."""
        paragraph = self.paragraph(paragraph_index)
        font = paragraph.Range.Font
        applied: dict[str, Any] = {}

        if font_name:
            font.Name = str(font_name)
            applied["font_name"] = str(font_name)
        if font_size is not None:
            font.Size = float(font_size)
            applied["font_size"] = float(font_size)
        if color is not None:
            font.Color = parse_color(color)
            applied["color"] = color
        if bold is not None:
            font.Bold = bool(bold)
            applied["bold"] = bool(bold)
        if italic is not None:
            font.Italic = bool(italic)
            applied["italic"] = bool(italic)
        if underline is not None:
            font.Underline = 1 if underline else 0
            applied["underline"] = bool(underline)

        return {"paragraph_index": int(paragraph_index), "applied": applied}

    @action("set_paragraph_alignment")
    def set_paragraph_alignment(
        self, paragraph_index: int, alignment: str
    ) -> dict[str, Any]:
        """Ustawia wyrownanie akapitu: left / center / right / justify."""
        key = str(alignment).strip().lower()
        if key not in WD_ALIGNMENTS:
            raise InvalidReferenceError(
                f"Nieznane wyrownanie '{alignment}'. Dostepne: "
                f"{', '.join(sorted(WD_ALIGNMENTS))}"
            )

        paragraph = self.paragraph(paragraph_index)
        paragraph.Alignment = WD_ALIGNMENTS[key]
        return {"paragraph_index": int(paragraph_index), "alignment": key}

    @action("apply_style")
    def apply_style(self, paragraph_index: int, style_name: str) -> dict[str, Any]:
        """Nadaje akapitowi styl (np. ``Heading 2``, ``Quote``, styl wlasny)."""
        paragraph = self.paragraph(paragraph_index)
        applied = self._apply_named_style(paragraph.Range, style_name)
        return {"paragraph_index": int(paragraph_index), "style": applied}

    @action("set_page_margins")
    def set_page_margins(
        self,
        top: float,
        bottom: float,
        left: float,
        right: float,
        unit: str = "pt",
    ) -> dict[str, Any]:
        """Ustawia marginesy strony; ``unit`` pozwala podac cm, mm, cale lub punkty."""
        document = self.document()
        setup = document.PageSetup

        values = {
            "top": points(top, unit),
            "bottom": points(bottom, unit),
            "left": points(left, unit),
            "right": points(right, unit),
        }

        setup.TopMargin = values["top"]
        setup.BottomMargin = values["bottom"]
        setup.LeftMargin = values["left"]
        setup.RightMargin = values["right"]

        return {"margins_points": {key: round(value, 2) for key, value in values.items()}}

    @action("insert_image")
    def insert_image(
        self,
        image_path: str,
        width: float | None = None,
        height: float | None = None,
        position: str = "inline",
    ) -> dict[str, Any]:
        """Wstawia obraz w tekscie (``inline``) albo jako obiekt plywajacy (``float``)."""
        target_path = self.resolve_existing_path(image_path)
        document = self.document()
        mode = str(position).strip().lower()

        if mode in ("inline", "w_tekscie"):
            shape = document.InlineShapes.AddPicture(
                FileName=target_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=self._end_range(document),
            )
        elif mode in ("float", "floating", "plywajacy"):
            shape = document.Shapes.AddPicture(
                FileName=target_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=50,
                Top=50,
            )
        else:
            raise InvalidReferenceError(
                f"Nieznana pozycja obrazu '{position}' - uzyj 'inline' albo 'float'"
            )

        if width is not None:
            shape.Width = float(width)
        if height is not None:
            shape.Height = float(height)

        return {
            "image_path": target_path,
            "position": mode,
            "width": round(float(shape.Width), 2),
            "height": round(float(shape.Height), 2),
        }

    @action("insert_table")
    def insert_table(
        self,
        rows: int,
        cols: int,
        data: list[list[Any]] | None = None,
        position: int | None = None,
        header_bold: bool = True,
    ) -> dict[str, Any]:
        """Wstawia tabele na koncu dokumentu albo po wskazanym akapicie."""
        row_count, column_count = int(rows), int(cols)
        if row_count < 1 or column_count < 1:
            raise InvalidReferenceError("Tabela musi miec co najmniej 1 wiersz i 1 kolumne")

        document = self.document()

        if position is None:
            target = self._end_range(document)
        else:
            paragraph = self.paragraph(position)
            target = paragraph.Range
            target.Collapse(WD_COLLAPSE_END)

        table = document.Tables.Add(target, row_count, column_count)

        try:
            table.Borders.InsideLineStyle = WD_LINE_STYLE_SINGLE
            table.Borders.OutsideLineStyle = WD_LINE_STYLE_SINGLE
        except com_error:
            pass

        filled = 0
        for row_index, row in enumerate(data or [], start=1):
            if row_index > row_count:
                break
            for column_index, value in enumerate(row, start=1):
                if column_index > column_count:
                    break
                if value is None:
                    continue
                cell = table.Cell(row_index, column_index)
                cell.Range.Text = str(value)
                if header_bold and row_index == 1:
                    cell.Range.Font.Bold = True
                filled += 1

        return {
            "rows": row_count,
            "cols": column_count,
            "cells_filled": filled,
            "position": position,
        }

    @action("insert_header")
    def insert_header(self, text: str, section: int = 1) -> dict[str, Any]:
        """Ustawia tekst naglowka strony w wybranej sekcji."""
        document = self.document()
        index = self.require_index(section, document.Sections.Count, "section")
        header = document.Sections(index).Headers(WD_HEADER_FOOTER_PRIMARY)
        header.Range.Text = str(text)
        return {"section": index, "header": str(text)}

    @action("insert_footer")
    def insert_footer(self, text: str, section: int = 1) -> dict[str, Any]:
        """Ustawia tekst stopki w wybranej sekcji."""
        document = self.document()
        index = self.require_index(section, document.Sections.Count, "section")
        footer = document.Sections(index).Footers(WD_HEADER_FOOTER_PRIMARY)
        footer.Range.Text = str(text)
        return {"section": index, "footer": str(text)}

    @action("add_page_numbers")
    def add_page_numbers(
        self, alignment: str = "center", first_page: bool = True, section: int = 1
    ) -> dict[str, Any]:
        """Wstawia numery stron w stopce."""
        key = str(alignment).strip().lower()
        if key not in WD_ALIGNMENTS:
            raise InvalidReferenceError(
                f"Nieznane wyrownanie '{alignment}'. Dostepne: "
                f"{', '.join(sorted(WD_ALIGNMENTS))}"
            )

        document = self.document()
        index = self.require_index(section, document.Sections.Count, "section")
        footer = document.Sections(index).Footers(WD_HEADER_FOOTER_PRIMARY)
        footer.PageNumbers.Add(
            PageNumberAlignment=WD_ALIGNMENTS[key], FirstPage=bool(first_page)
        )

        return {"section": index, "alignment": key, "first_page": bool(first_page)}

    @action("insert_table_of_contents")
    def insert_table_of_contents(
        self, levels: int = 3, position: str = "start"
    ) -> dict[str, Any]:
        """Wstawia spis tresci zbudowany ze stylow naglowkow."""
        depth = max(1, min(int(levels), 9))
        document = self.document()

        if str(position).strip().lower() in ("start", "poczatek"):
            target = document.Range(0, 0)
        else:
            target = self._end_range(document)

        try:
            toc = document.TablesOfContents.Add(
                Range=target,
                UseHeadingStyles=True,
                UpperHeadingLevel=1,
                LowerHeadingLevel=depth,
            )
            toc.Update()
        except com_error as exc:
            if is_connection_error(exc):
                raise
            raise UnsupportedOperationError(
                "Nie udalo sie wstawic spisu tresci - dokument musi zawierac "
                "akapity ze stylami naglowkow"
            ) from exc

        return {"levels": depth, "position": str(position).lower()}


__all__ = ["WordController"]
