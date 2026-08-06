"""Kontroler PowerPointa - tworzenie i edycja prezentacji przez COM.

Wszystkie akcje pracuja na aktywnej prezentacji (ostatnio otwartej lub
utworzonej). Indeksy slajdow sa 1-based, tak jak w samym PowerPoincie,
a pozycje i rozmiary podaje sie w punktach (1 cm = 28.35 pt).
"""

from __future__ import annotations

import os
from typing import Any

from bridge.controllers.base import BaseController, action
from bridge.utils.com_helpers import (
    CHART_TYPES,
    PP_LAYOUTS,
    PP_SAVE_FORMATS,
    SHAPE_TYPES,
    bgr_to_hex,
    com_error,
    lookup_constant,
    normalize_path,
    parse_color,
    save_format_for,
    to_python,
)
from bridge.utils.errors import (
    DocumentNotFoundError,
    InvalidReferenceError,
    UnsupportedOperationError,
)

MSO_TEXT_HORIZONTAL = 1
MSO_TRUE = -1
MSO_FALSE = 0

PLACEHOLDER_TYPES = {
    1: "title",
    2: "body",
    3: "center_title",
    4: "subtitle",
    5: "vertical_title",
    6: "vertical_body",
    7: "object",
    8: "chart",
    9: "bitmap",
    10: "media_clip",
    11: "org_chart",
    12: "table",
    13: "slide_number",
    14: "header",
    15: "footer",
    16: "date",
}

CONTENT_PLACEHOLDERS = (2, 4, 6, 7, 8, 12)
TITLE_PLACEHOLDERS = (1, 3, 5)

THEME_DIRECTORIES = (
    r"C:\Program Files\Microsoft Office\root\Document Themes 16",
    r"C:\Program Files (x86)\Microsoft Office\root\Document Themes 16",
    r"C:\Program Files\Microsoft Office\Document Themes 16",
)


class PowerPointController(BaseController):
    """Akcje ``ppt_*`` - operacje na zywej instancji PowerPointa."""

    APP_KEY = "powerpoint"
    DISPLAY_NAME = "PowerPoint"

    def presentation(self) -> Any:
        """Aktywna prezentacja albo czytelny blad, gdy nic nie jest otwarte."""
        app = self.app
        if app.Presentations.Count == 0:
            raise DocumentNotFoundError(
                "Brak otwartej prezentacji - uzyj ppt_create_presentation albo "
                "ppt_open_presentation"
            )
        try:
            return app.ActivePresentation
        except com_error:
            return app.Presentations(app.Presentations.Count)

    def slide(self, slide_index: Any) -> Any:
        """Slajd o zadanym indeksie (1-based) z walidacja zakresu."""
        presentation = self.presentation()
        index = self.require_index(slide_index, presentation.Slides.Count, "slide_index")
        return presentation.Slides(index)

    def _goto_slide(self, index: int) -> None:
        """Przewija okno PowerPointa na slajd, zeby uzytkownik widzial zmiane."""
        try:
            self.presentation().Windows(1).View.GotoSlide(index)
        except com_error:
            pass

    def _presentation_summary(self, presentation: Any) -> dict[str, Any]:
        return {
            "name": to_python(presentation.Name),
            "path": to_python(presentation.FullName) if presentation.Path else None,
            "slide_count": int(presentation.Slides.Count),
            "saved": bool(presentation.Saved),
        }

    def _shape_summary(self, shape: Any) -> dict[str, Any]:
        info: dict[str, Any] = {
            "shape_id": int(shape.Id),
            "name": to_python(shape.Name),
            "left": round(float(shape.Left), 2),
            "top": round(float(shape.Top), 2),
            "width": round(float(shape.Width), 2),
            "height": round(float(shape.Height), 2),
            "has_text": False,
            "text": None,
            "placeholder": None,
        }

        try:
            if shape.Type == 14:
                info["placeholder"] = PLACEHOLDER_TYPES.get(
                    int(shape.PlaceholderFormat.Type), "unknown"
                )
        except com_error:
            pass

        try:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                info["has_text"] = True
                info["text"] = to_python(shape.TextFrame.TextRange.Text)
        except com_error:
            pass

        try:
            if shape.HasTable:
                info["table"] = {
                    "rows": int(shape.Table.Rows.Count),
                    "columns": int(shape.Table.Columns.Count),
                }
        except com_error:
            pass

        return info

    def _find_shape(self, slide: Any, shape_id: Any) -> Any:
        """Znajduje ksztalt po ``Id`` (liczba) albo po nazwie (tekst)."""
        if isinstance(shape_id, str) and not shape_id.isdigit():
            for index in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(index)
                if str(shape.Name).lower() == shape_id.lower():
                    return shape
            raise InvalidReferenceError(f"Slajd nie zawiera ksztaltu '{shape_id}'")

        wanted = int(shape_id)
        for index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(index)
            if int(shape.Id) == wanted:
                return shape
        raise InvalidReferenceError(f"Slajd nie zawiera ksztaltu o id {wanted}")

    def _title_shape(self, slide: Any) -> Any | None:
        try:
            if slide.Shapes.HasTitle:
                return slide.Shapes.Title
        except com_error:
            pass
        return None

    def _placeholder_frame(self, slide: Any, placeholder: Any) -> Any:
        """Zwraca ``TextFrame`` wskazanego placeholdera (``content`` / ``title``)."""
        if isinstance(placeholder, (int, str)) and str(placeholder).isdigit():
            return self._find_shape(slide, int(placeholder)).TextFrame

        wanted = str(placeholder or "content").strip().lower()
        if wanted in ("title", "tytul"):
            shape = self._title_shape(slide)
            if shape is None:
                raise InvalidReferenceError("Slajd nie ma placeholdera tytulu")
            return shape.TextFrame

        wanted_types = CONTENT_PLACEHOLDERS if wanted in ("content", "body", "tresc") else None
        for index in range(1, slide.Shapes.Placeholders.Count + 1):
            shape = slide.Shapes.Placeholders(index)
            try:
                placeholder_type = int(shape.PlaceholderFormat.Type)
            except com_error:
                continue
            if wanted_types is None or placeholder_type in wanted_types:
                if shape.HasTextFrame:
                    return shape.TextFrame

        if wanted_types is None:
            raise InvalidReferenceError(f"Nieznany placeholder: {placeholder!r}")

        shape = slide.Shapes.AddTextbox(
            MSO_TEXT_HORIZONTAL, 60, 140, 600, 300
        )
        return shape.TextFrame

    @action("create_presentation")
    def create_presentation(self, path: str, template: str | None = None) -> dict[str, Any]:
        """Tworzy nowa prezentacje (opcjonalnie z szablonu .potx/.thmx) i zapisuje ja."""
        target = self.resolve_target_path(path)
        presentation = self.app.Presentations.Add(WithWindow=MSO_TRUE)

        if template:
            presentation.ApplyTemplate(self.resolve_existing_path(template))

        presentation.SaveAs(
            target, save_format_for(target, PP_SAVE_FORMATS, PP_SAVE_FORMATS[".pptx"])
        )
        return self._presentation_summary(presentation)

    @action("open_presentation")
    def open_presentation(self, path: str) -> dict[str, Any]:
        """Otwiera plik; jesli jest juz otwarty, tylko aktywuje jego okno."""
        target = self.resolve_existing_path(path)
        app = self.app

        for index in range(1, app.Presentations.Count + 1):
            presentation = app.Presentations(index)
            if os.path.normcase(str(presentation.FullName)) == os.path.normcase(target):
                try:
                    presentation.Windows(1).Activate()
                except com_error:
                    pass
                return {**self._presentation_summary(presentation), "already_open": True}

        presentation = app.Presentations.Open(target, ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        return {**self._presentation_summary(presentation), "already_open": False}

    @action("save")
    def save(self, path: str | None = None) -> dict[str, Any]:
        """Zapisuje prezentacje (``Save``) albo zapisuje jako nowy plik (``SaveAs``)."""
        presentation = self.presentation()

        if path:
            target = self.resolve_target_path(path)
            presentation.SaveAs(
                target, save_format_for(target, PP_SAVE_FORMATS, PP_SAVE_FORMATS[".pptx"])
            )
        elif not presentation.Path:
            raise InvalidReferenceError(
                "Prezentacja nie ma jeszcze pliku - podaj parametr path"
            )
        else:
            presentation.Save()

        return self._presentation_summary(presentation)

    @action("close")
    def close(self, save: bool = True) -> dict[str, Any]:
        """Zamyka prezentacje, opcjonalnie zapisujac zmiany."""
        presentation = self.presentation()
        name = str(presentation.Name)

        if save:
            if not presentation.Path:
                raise InvalidReferenceError(
                    "Prezentacja nie byla zapisana - najpierw ppt_save z parametrem path"
                )
            presentation.Save()
        else:
            presentation.Saved = MSO_TRUE

        presentation.Close()
        return {"closed": name, "saved": bool(save)}

    @action("get_presentation_info")
    def get_presentation_info(self) -> dict[str, Any]:
        """Podstawowe metadane prezentacji: rozmiar slajdu, motyw, sciezka."""
        presentation = self.presentation()
        info = self._presentation_summary(presentation)

        info["slide_width"] = round(float(presentation.PageSetup.SlideWidth), 2)
        info["slide_height"] = round(float(presentation.PageSetup.SlideHeight), 2)

        try:
            info["theme"] = to_python(presentation.Designs(1).Name)
        except com_error:
            info["theme"] = None

        try:
            info["read_only"] = bool(presentation.ReadOnly)
        except com_error:
            info["read_only"] = False

        return info

    @action("list_slides")
    def list_slides(self) -> dict[str, Any]:
        """Lista slajdow: indeks, tytul, uklad i liczba ksztaltow."""
        presentation = self.presentation()
        slides = []

        for index in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(index)
            title_shape = self._title_shape(slide)
            title = None
            if title_shape is not None and title_shape.TextFrame.HasText:
                title = to_python(title_shape.TextFrame.TextRange.Text)

            try:
                layout = to_python(slide.CustomLayout.Name)
            except com_error:
                layout = to_python(slide.Layout)

            slides.append(
                {
                    "index": index,
                    "title": title,
                    "layout": layout,
                    "shape_count": int(slide.Shapes.Count),
                }
            )

        return {"slide_count": len(slides), "slides": slides}

    @action("get_slide_content")
    def get_slide_content(self, slide_index: int) -> dict[str, Any]:
        """Pelna zawartosc slajdu: ksztalty, ich pozycje, teksty i notatki."""
        slide = self.slide(slide_index)
        shapes = [
            self._shape_summary(slide.Shapes(index))
            for index in range(1, slide.Shapes.Count + 1)
        ]

        notes = None
        try:
            notes_frame = slide.NotesPage.Shapes.Placeholders(2).TextFrame
            if notes_frame.HasText:
                notes = to_python(notes_frame.TextRange.Text)
        except com_error:
            notes = None

        try:
            layout = to_python(slide.CustomLayout.Name)
        except com_error:
            layout = to_python(slide.Layout)

        return {
            "slide_index": int(slide_index),
            "layout": layout,
            "shapes": shapes,
            "notes": notes,
        }

    @action("add_slide")
    def add_slide(
        self,
        layout: str = "title_content",
        index: int | None = None,
        title: str | None = None,
    ) -> dict[str, Any]:
        """Dodaje slajd o wskazanym ukladzie; ``index=None`` oznacza koniec."""
        presentation = self.presentation()
        count = presentation.Slides.Count
        position = count + 1 if index is None else max(1, min(int(index), count + 1))
        layout_constant = lookup_constant(layout, PP_LAYOUTS, "layout")

        slide = presentation.Slides.Add(position, layout_constant)

        if title:
            title_shape = self._title_shape(slide)
            if title_shape is None:
                title_shape = slide.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, 60, 40, 600, 60)
            title_shape.TextFrame.TextRange.Text = title

        self._goto_slide(position)
        return {
            "slide_index": position,
            "layout": layout,
            "slide_count": int(presentation.Slides.Count),
        }

    @action("delete_slide")
    def delete_slide(self, slide_index: int) -> dict[str, Any]:
        """Usuwa slajd o podanym indeksie."""
        presentation = self.presentation()
        index = self.require_index(slide_index, presentation.Slides.Count, "slide_index")
        presentation.Slides(index).Delete()
        return {"deleted": index, "slide_count": int(presentation.Slides.Count)}

    @action("duplicate_slide")
    def duplicate_slide(self, slide_index: int) -> dict[str, Any]:
        """Duplikuje slajd - kopia laduje bezposrednio za oryginalem."""
        slide = self.slide(slide_index)
        copy = slide.Duplicate()
        try:
            new_index = int(copy.SlideIndex)
        except (com_error, AttributeError, TypeError):
            new_index = int(copy(1).SlideIndex)
        self._goto_slide(new_index)
        return {"source_index": int(slide_index), "slide_index": new_index}

    @action("reorder_slide")
    def reorder_slide(self, from_index: int, to_index: int) -> dict[str, Any]:
        """Przenosi slajd na inna pozycje."""
        presentation = self.presentation()
        count = presentation.Slides.Count
        source = self.require_index(from_index, count, "from_index")
        target = self.require_index(to_index, count, "to_index")
        presentation.Slides(source).MoveTo(target)
        self._goto_slide(target)
        return {"from_index": source, "to_index": target}

    @action("set_title")
    def set_title(self, slide_index: int, text: str) -> dict[str, Any]:
        """Ustawia tytul slajdu; gdy uklad go nie ma, wstawia pole tekstowe."""
        slide = self.slide(slide_index)
        shape = self._title_shape(slide)
        created = False

        if shape is None:
            shape = slide.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, 60, 40, 600, 60)
            shape.TextFrame.TextRange.Font.Size = 32
            created = True

        shape.TextFrame.TextRange.Text = text
        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "created_textbox": created,
        }

    @action("add_textbox")
    def add_textbox(
        self,
        slide_index: int,
        text: str,
        left: float,
        top: float,
        width: float,
        height: float,
        font_size: float | None = None,
        bold: bool = False,
        color: Any = None,
        align: str | None = None,
    ) -> dict[str, Any]:
        """Wstawia pole tekstowe w podanym miejscu slajdu (wspolrzedne w punktach)."""
        slide = self.slide(slide_index)
        shape = slide.Shapes.AddTextbox(
            MSO_TEXT_HORIZONTAL, float(left), float(top), float(width), float(height)
        )
        text_range = shape.TextFrame.TextRange
        text_range.Text = text

        if font_size is not None:
            text_range.Font.Size = float(font_size)
        if bold:
            text_range.Font.Bold = MSO_TRUE
        if color is not None:
            text_range.Font.Color.RGB = parse_color(color)
        if align:
            text_range.ParagraphFormat.Alignment = {
                "left": 1,
                "center": 2,
                "right": 3,
                "justify": 4,
            }.get(str(align).lower(), 1)

        self._goto_slide(int(slide_index))
        return {"slide_index": int(slide_index), "shape_id": int(shape.Id)}

    @action("add_bullet_list")
    def add_bullet_list(
        self,
        slide_index: int,
        items: list[Any],
        placeholder: Any = "content",
    ) -> dict[str, Any]:
        """Wypelnia placeholder lista punktowana z obsluga zagniezdzen.

        ``items`` przyjmuje teksty (``"Punkt"``) lub slowniki z poziomem
        wciecia (``{"text": "Podpunkt", "level": 2}``).
        """
        if not isinstance(items, list) or not items:
            raise InvalidReferenceError("Lista 'items' nie moze byc pusta")

        slide = self.slide(slide_index)
        frame = self._placeholder_frame(slide, placeholder)

        entries: list[tuple[str, int]] = []
        for item in items:
            if isinstance(item, dict):
                text = str(item.get("text", ""))
                level = int(item.get("level", 1))
            elif isinstance(item, (list, tuple)) and len(item) == 2:
                text, level = str(item[0]), int(item[1])
            else:
                text, level = str(item), 1
            entries.append((text, max(1, min(level, 5))))

        text_range = frame.TextRange
        text_range.Text = "\r".join(text for text, _ in entries)

        for position, (_, level) in enumerate(entries, start=1):
            try:
                text_range.Paragraphs(position).IndentLevel = level
            except com_error:
                pass

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(frame.Parent.Id),
            "items": len(entries),
        }

    @action("find_replace_text")
    def find_replace_text(
        self,
        old_text: str,
        new_text: str,
        slide_index: int | None = None,
        match_case: bool = False,
    ) -> dict[str, Any]:
        """Podmienia tekst na jednym slajdzie albo w calej prezentacji."""
        if not old_text:
            raise InvalidReferenceError("Parametr old_text nie moze byc pusty")

        presentation = self.presentation()
        if slide_index is None:
            indexes = list(range(1, presentation.Slides.Count + 1))
        else:
            indexes = [
                self.require_index(slide_index, presentation.Slides.Count, "slide_index")
            ]

        replacements = 0
        for index in indexes:
            slide = presentation.Slides(index)
            for shape_index in range(1, slide.Shapes.Count + 1):
                replacements += self._replace_in_shape(
                    slide.Shapes(shape_index), old_text, new_text, match_case
                )

        return {
            "replacements": replacements,
            "slides_scanned": len(indexes),
            "old_text": old_text,
            "new_text": new_text,
        }

    def _replace_in_shape(
        self, shape: Any, old_text: str, new_text: str, match_case: bool
    ) -> int:
        count = 0

        try:
            if shape.HasTable:
                table = shape.Table
                for row in range(1, table.Rows.Count + 1):
                    for column in range(1, table.Columns.Count + 1):
                        cell_frame = table.Cell(row, column).Shape.TextFrame
                        count += self._replace_in_range(
                            cell_frame, old_text, new_text, match_case
                        )
                return count
        except com_error:
            pass

        try:
            if shape.Type == 6:
                for index in range(1, shape.GroupItems.Count + 1):
                    count += self._replace_in_shape(
                        shape.GroupItems(index), old_text, new_text, match_case
                    )
                return count
        except com_error:
            pass

        try:
            if shape.HasTextFrame:
                count += self._replace_in_range(shape.TextFrame, old_text, new_text, match_case)
        except com_error:
            pass

        return count

    def _replace_in_range(
        self, text_frame: Any, old_text: str, new_text: str, match_case: bool
    ) -> int:
        if not text_frame.HasText:
            return 0

        text_range = text_frame.TextRange
        count = 0
        after = 0

        for _ in range(500):
            found = text_range.Replace(
                FindWhat=old_text,
                ReplaceWhat=new_text,
                After=after,
                MatchCase=MSO_TRUE if match_case else MSO_FALSE,
                WholeWords=MSO_FALSE,
            )
            if not found:
                break
            count += 1
            after = int(found.Start) + len(new_text) - 1

        return count

    @action("set_speaker_notes")
    def set_speaker_notes(self, slide_index: int, text: str) -> dict[str, Any]:
        """Ustawia notatki prelegenta dla slajdu."""
        slide = self.slide(slide_index)
        try:
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text
        except com_error as exc:
            raise UnsupportedOperationError(
                "Slajd nie ma miejsca na notatki prelegenta"
            ) from exc
        return {"slide_index": int(slide_index), "characters": len(text)}

    @action("set_text_style")
    def set_text_style(
        self,
        slide_index: int,
        shape_id: Any,
        font_name: str | None = None,
        font_size: float | None = None,
        color: Any = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
    ) -> dict[str, Any]:
        """Formatuje caly tekst wskazanego ksztaltu."""
        slide = self.slide(slide_index)
        shape = self._find_shape(slide, shape_id)

        if not shape.HasTextFrame:
            raise InvalidReferenceError("Wskazany ksztalt nie zawiera tekstu")

        font = shape.TextFrame.TextRange.Font
        applied: dict[str, Any] = {}

        if font_name:
            font.Name = font_name
            applied["font_name"] = font_name
        if font_size is not None:
            font.Size = float(font_size)
            applied["font_size"] = float(font_size)
        if color is not None:
            font.Color.RGB = parse_color(color)
            applied["color"] = color
        if bold is not None:
            font.Bold = MSO_TRUE if bold else MSO_FALSE
            applied["bold"] = bool(bold)
        if italic is not None:
            font.Italic = MSO_TRUE if italic else MSO_FALSE
            applied["italic"] = bool(italic)
        if underline is not None:
            font.Underline = MSO_TRUE if underline else MSO_FALSE
            applied["underline"] = bool(underline)

        self._goto_slide(int(slide_index))
        return {"slide_index": int(slide_index), "shape_id": int(shape.Id), "applied": applied}

    @action("apply_theme")
    def apply_theme(self, theme_name_or_path: str) -> dict[str, Any]:
        """Nadaje motyw z pliku ``.thmx``/``.potx`` albo z galerii motywow Office."""
        theme_path = self._resolve_theme(theme_name_or_path)
        presentation = self.presentation()
        presentation.ApplyTemplate(theme_path)
        return {"theme": os.path.basename(theme_path), "path": theme_path}

    def _resolve_theme(self, name_or_path: str) -> str:
        candidate = normalize_path(name_or_path)
        if os.path.isfile(candidate):
            return candidate

        stem = os.path.splitext(os.path.basename(str(name_or_path)))[0].lower()
        for directory in THEME_DIRECTORIES:
            if not os.path.isdir(directory):
                continue
            for entry in os.listdir(directory):
                if os.path.splitext(entry)[0].lower() == stem and entry.lower().endswith(
                    (".thmx", ".potx")
                ):
                    return os.path.join(directory, entry)

        raise DocumentNotFoundError(
            f"Nie znaleziono motywu '{name_or_path}'. COM przyjmuje sciezke do pliku "
            ".thmx lub .potx - podaj pelna sciezke albo nazwe motywu z galerii Office."
        )

    @action("set_background")
    def set_background(
        self,
        slide_index: int,
        color: Any = None,
        image_path: str | None = None,
    ) -> dict[str, Any]:
        """Ustawia tlo slajdu - jednolity kolor albo obraz."""
        if color is None and not image_path:
            raise InvalidReferenceError("Podaj kolor albo sciezke do obrazu tla")

        slide = self.slide(slide_index)
        slide.FollowMasterBackground = MSO_FALSE
        fill = slide.Background.Fill

        if image_path:
            fill.UserPicture(self.resolve_existing_path(image_path))
            applied = {"image_path": self.resolve_existing_path(image_path)}
        else:
            fill.Solid()
            fill.ForeColor.RGB = parse_color(color)
            applied = {"color": bgr_to_hex(fill.ForeColor.RGB)}

        self._goto_slide(int(slide_index))
        return {"slide_index": int(slide_index), **applied}

    @action("set_slide_layout")
    def set_slide_layout(self, slide_index: int, layout_name: str) -> dict[str, Any]:
        """Zmienia uklad slajdu - po nazwie ukladu z wzorca albo nazwie standardowej."""
        slide = self.slide(slide_index)
        presentation = self.presentation()

        wanted = str(layout_name).strip().lower()
        layouts = presentation.SlideMaster.CustomLayouts
        for index in range(1, layouts.Count + 1):
            if str(layouts(index).Name).strip().lower() == wanted:
                slide.CustomLayout = layouts(index)
                self._goto_slide(int(slide_index))
                return {
                    "slide_index": int(slide_index),
                    "layout": to_python(layouts(index).Name),
                    "source": "custom_layout",
                }

        slide.Layout = lookup_constant(layout_name, PP_LAYOUTS, "layout_name")
        self._goto_slide(int(slide_index))
        return {"slide_index": int(slide_index), "layout": layout_name, "source": "builtin"}

    @action("add_image")
    def add_image(
        self,
        slide_index: int,
        image_path: str,
        left: float,
        top: float,
        width: float | None = None,
        height: float | None = None,
    ) -> dict[str, Any]:
        """Wstawia obraz; brak width/height zachowuje oryginalne proporcje."""
        slide = self.slide(slide_index)
        picture = slide.Shapes.AddPicture(
            FileName=self.resolve_existing_path(image_path),
            LinkToFile=MSO_FALSE,
            SaveWithDocument=MSO_TRUE,
            Left=float(left),
            Top=float(top),
            Width=float(width) if width is not None else -1,
            Height=float(height) if height is not None else -1,
        )

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(picture.Id),
            "width": round(float(picture.Width), 2),
            "height": round(float(picture.Height), 2),
        }

    @action("add_chart")
    def add_chart(
        self,
        slide_index: int,
        chart_type: str,
        categories: list[Any],
        series_data: Any,
        left: float,
        top: float,
        width: float,
        height: float,
        title: str | None = None,
    ) -> dict[str, Any]:
        """Wstawia wykres i wypelnia jego arkusz danymi.

        ``series_data`` przyjmuje slownik ``{"nazwa": [wartosci]}``, liste
        slownikow ``{"name": ..., "values": [...]}`` albo sama liste serii.
        """
        if not categories:
            raise InvalidReferenceError("Lista 'categories' nie moze byc pusta")

        slide = self.slide(slide_index)
        chart_constant = lookup_constant(chart_type, CHART_TYPES, "chart_type")
        series = _normalize_series(series_data, len(categories))

        try:
            shape = slide.Shapes.AddChart2(
                -1, chart_constant, float(left), float(top), float(width), float(height)
            )
        except (com_error, AttributeError):
            shape = slide.Shapes.AddChart(
                chart_constant, float(left), float(top), float(width), float(height)
            )

        chart = shape.Chart
        self._fill_chart_data(chart, categories, series)

        if title:
            chart.HasTitle = MSO_TRUE
            chart.ChartTitle.Text = title

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "chart_type": chart_type,
            "series": [name for name, _ in series],
            "categories": len(categories),
        }

    def _fill_chart_data(
        self, chart: Any, categories: list[Any], series: list[tuple[str, list[Any]]]
    ) -> None:
        """Wpisuje dane do arkusza osadzonego w wykresie i ustawia zakres zrodlowy."""
        chart.ChartData.Activate()
        workbook = chart.ChartData.Workbook
        worksheet = workbook.Worksheets(1)

        try:
            while worksheet.ListObjects.Count:
                worksheet.ListObjects(1).Unlist()
        except com_error:
            pass

        worksheet.Cells.Clear()
        worksheet.Cells(1, 1).Value = ""

        for column, (name, _values) in enumerate(series, start=2):
            worksheet.Cells(1, column).Value = name

        for row, category in enumerate(categories, start=2):
            worksheet.Cells(row, 1).Value = category

        for column, (_name, values) in enumerate(series, start=2):
            for row, value in enumerate(values, start=2):
                worksheet.Cells(row, column).Value = value

        last_row = 1 + len(categories)
        last_column = 1 + len(series)
        data_range = worksheet.Range(
            worksheet.Cells(1, 1), worksheet.Cells(last_row, last_column)
        )
        chart.SetSourceData(f"='{worksheet.Name}'!{data_range.Address(True, True)}")

        try:
            workbook.Close()
        except com_error:
            pass

    @action("add_table")
    def add_table(
        self,
        slide_index: int,
        rows: int,
        cols: int,
        data: list[list[Any]] | None,
        left: float,
        top: float,
        width: float,
        height: float,
        header_bold: bool = True,
    ) -> dict[str, Any]:
        """Wstawia tabele i wypelnia ja danymi (nadmiarowe komorki sa pomijane)."""
        rows, cols = int(rows), int(cols)
        if rows < 1 or cols < 1:
            raise InvalidReferenceError("Tabela musi miec co najmniej 1 wiersz i 1 kolumne")

        slide = self.slide(slide_index)
        shape = slide.Shapes.AddTable(
            rows, cols, float(left), float(top), float(width), float(height)
        )
        table = shape.Table
        filled = 0

        for row_index, row in enumerate(data or [], start=1):
            if row_index > rows:
                break
            for column_index, value in enumerate(row, start=1):
                if column_index > cols:
                    break
                if value is None:
                    continue
                cell = table.Cell(row_index, column_index)
                cell.Shape.TextFrame.TextRange.Text = str(value)
                if header_bold and row_index == 1:
                    cell.Shape.TextFrame.TextRange.Font.Bold = MSO_TRUE
                filled += 1

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "rows": rows,
            "cols": cols,
            "cells_filled": filled,
        }

    @action("add_shape")
    def add_shape(
        self,
        slide_index: int,
        shape_type: str,
        left: float,
        top: float,
        width: float,
        height: float,
        fill_color: Any = None,
        text: str | None = None,
    ) -> dict[str, Any]:
        """Wstawia ksztalt (prostokat, strzalka, gwiazda...) z opcjonalnym tekstem."""
        slide = self.slide(slide_index)
        shape_constant = lookup_constant(shape_type, SHAPE_TYPES, "shape_type")
        shape = slide.Shapes.AddShape(
            shape_constant, float(left), float(top), float(width), float(height)
        )

        if fill_color is not None:
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = parse_color(fill_color)
        if text:
            shape.TextFrame.TextRange.Text = text

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "shape_type": shape_type,
        }


def _normalize_series(series_data: Any, expected_length: int) -> list[tuple[str, list[Any]]]:
    """Sprowadza rozne formaty serii danych do listy par ``(nazwa, wartosci)``."""
    series: list[tuple[str, list[Any]]] = []

    if isinstance(series_data, dict):
        for name, values in series_data.items():
            series.append((str(name), list(values)))
    elif isinstance(series_data, (list, tuple)):
        for position, entry in enumerate(series_data, start=1):
            if isinstance(entry, dict):
                name = str(entry.get("name", f"Seria {position}"))
                values = list(entry.get("values", []))
            elif isinstance(entry, (list, tuple)):
                name, values = f"Seria {position}", list(entry)
            else:
                name, values = f"Seria {position}", [entry]
            series.append((name, values))
    else:
        raise InvalidReferenceError(
            "series_data musi byc slownikiem, lista serii albo lista list wartosci"
        )

    if not series:
        raise InvalidReferenceError("Brak danych do wykresu")

    normalized = []
    for name, values in series:
        padded = list(values) + [None] * (expected_length - len(values))
        normalized.append((name, padded[:expected_length]))
    return normalized


__all__ = ["PowerPointController"]
