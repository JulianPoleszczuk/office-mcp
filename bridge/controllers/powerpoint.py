"""PowerPoint controller - building and editing presentations over COM.

Every action works on the presentation the plugin is pinned to (the last one
opened or created). Slide indexes are 1-based, as in PowerPoint itself, and
positions and sizes are given in points (1 cm = 28.35 pt).
"""

from __future__ import annotations

import os
import time
from typing import Any

from bridge.controllers.base import BaseController, action
from bridge.utils.com_helpers import (
    CHART_TYPES,
    MSO_ALIGN,
    MSO_ANCHORS,
    MSO_ANIM_EFFECTS,
    MSO_ANIM_LEVELS,
    MSO_ANIM_TRIGGERS,
    MSO_DISTRIBUTE,
    MSO_GRADIENT_STYLES,
    MSO_LINE_DASHES,
    MSO_THEME_COLORS,
    MSO_ZORDER,
    PP_EXPORT_FILTERS,
    PP_LAYOUTS,
    PP_SAVE_FORMATS,
    PP_TRANSITIONS,
    SHAPE_TYPES,
    apply_chart_format,
    bgr_to_hex,
    com_address,
    com_error,
    constant_name,
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

MSO_THEME_LATIN = 1
MSO_SHADOW_OUTER = 2
PP_MOUSE_CLICK = 1
PP_ACTION_HYPERLINK = 7
PP_ACTION_NONE = 0
MSO_ANIM_MEDIA_PLAY = 83
MSO_SMARTART_NODE_BELOW = 5
PP_AUTOSIZE_NONE = 0
PP_AUTOSIZE_FIT = 1

PP_ALIGNMENTS: dict[str, int] = {
    "left": 1,
    "center": 2,
        "right": 3,
    "justify": 4,
}

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

# Name of the text box that set_title creates on a layout without a title
# placeholder. Thanks to it the "title" shortcut works in every other tool,
# not only in the one that inserted the title.
TITLE_FALLBACK_NAME = "office-mcp Title"

CONTENT_PLACEHOLDERS = (2, 4, 6, 7, 8, 12)
TITLE_PLACEHOLDERS = (1, 3, 5)

THEME_DIRECTORIES = (
    r"C:\Program Files\Microsoft Office\root\Document Themes 16",
    r"C:\Program Files (x86)\Microsoft Office\root\Document Themes 16",
    r"C:\Program Files\Microsoft Office\Document Themes 16",
)


class PowerPointController(BaseController):
    """``ppt_*`` actions - operations on a live PowerPoint instance."""

    APP_KEY = "powerpoint"
    DISPLAY_NAME = "PowerPoint"
    ALERTS_OFF = 1

    def __init__(self, connection: Any) -> None:
        super().__init__(connection)
        # Path of the presentation we work on. PowerPoint ignores
        # Windows.Activate() when the app is not in the foreground, so
        # ActivePresentation can point at a completely different file than the
        # one we just opened - and then every later tool silently edits the
        # wrong document.
        self._target_path: str | None = None

    def _remember(self, presentation: Any) -> Any:
        """Pins a presentation as the current one for later calls."""
        try:
            self._target_path = (
                os.path.normcase(str(presentation.FullName))
                if presentation.Path
                else None
            )
        except com_error:
            self._target_path = None
        return presentation

    def presentation(self) -> Any:
        """The presentation we work on - pinned, not 'active'."""
        app = self.app
        if app.Presentations.Count == 0:
            raise DocumentNotFoundError(
                "No presentation open - use ppt_create_presentation or "
                "ppt_open_presentation"
            )

        if self._target_path:
            for index in range(1, app.Presentations.Count + 1):
                candidate = app.Presentations(index)
                try:
                    if candidate.Path and (
                        os.path.normcase(str(candidate.FullName)) == self._target_path
                    ):
                        return candidate
                except com_error:
                    continue
            # The pinned presentation was closed outside the plugin.
            self._target_path = None

        try:
            return app.ActivePresentation
        except com_error:
            return app.Presentations(app.Presentations.Count)

    def slide(self, slide_index: Any) -> Any:
        """Slide at the given 1-based index, with range validation."""
        presentation = self.presentation()
        index = self.require_index(slide_index, presentation.Slides.Count, "slide_index")
        return presentation.Slides(index)

    def _goto_slide(self, index: int) -> None:
        """Scrolls the PowerPoint window to the slide so the user sees the change."""
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
        """Finds a shape by ``Id`` (number) or by name (text)."""
        if isinstance(shape_id, str) and not shape_id.isdigit():
            for index in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(index)
                if str(shape.Name).lower() == shape_id.lower():
                    return shape
            raise InvalidReferenceError(f"The slide has no shape named '{shape_id}'")

        wanted = int(shape_id)
        for index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(index)
            if int(shape.Id) == wanted:
                return shape
        raise InvalidReferenceError(f"The slide has no shape with id {wanted}")

    def _title_shape(self, slide: Any) -> Any | None:
        try:
            if slide.Shapes.HasTitle:
                return slide.Shapes.Title
        except com_error:
            pass

        for index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(index)
            try:
                if str(shape.Name) == TITLE_FALLBACK_NAME:
                    return shape
            except com_error:
                continue
        return None

    def _placeholder_frame(self, slide: Any, placeholder: Any) -> Any:
        """Returns the ``TextFrame`` of the given placeholder (``content`` / ``title``)."""
        if isinstance(placeholder, (int, str)) and str(placeholder).isdigit():
            return self._find_shape(slide, int(placeholder)).TextFrame

        wanted = str(placeholder or "content").strip().lower()
        if wanted in ("title",):
            shape = self._title_shape(slide)
            if shape is None:
                raise InvalidReferenceError("The slide has no title placeholder")
            return shape.TextFrame

        wanted_types = CONTENT_PLACEHOLDERS if wanted in ("content", "body") else None
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
            raise InvalidReferenceError(f"Unknown placeholder: {placeholder!r}")

        shape = slide.Shapes.AddTextbox(
            MSO_TEXT_HORIZONTAL, 60, 140, 600, 300
        )
        return shape.TextFrame

    @action("create_presentation")
    def create_presentation(self, path: str, template: str | None = None) -> dict[str, Any]:
        """Creates a presentation (optionally from a .potx/.thmx template) and saves it."""
        target = self.resolve_target_path(path)
        presentation = self.app.Presentations.Add(WithWindow=MSO_TRUE)

        if template:
            presentation.ApplyTemplate(self.resolve_existing_path(template))

        with self.alerts_suppressed():
            presentation.SaveAs(
                target, save_format_for(target, PP_SAVE_FORMATS, PP_SAVE_FORMATS[".pptx"])
            )
        return self._presentation_summary(self._remember(presentation))

    @action("open_presentation")
    def open_presentation(self, path: str) -> dict[str, Any]:
        """Opens the file; if it is already open, just activates its window."""
        target = self.resolve_existing_path(path)
        app = self.app

        for index in range(1, app.Presentations.Count + 1):
            presentation = app.Presentations(index)
            if os.path.normcase(str(presentation.FullName)) == os.path.normcase(target):
                try:
                    presentation.Windows(1).Activate()
                except com_error:
                    pass
                return {
                    **self._presentation_summary(self._remember(presentation)),
                    "already_open": True,
                }

        presentation = app.Presentations.Open(target, ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        return {
            **self._presentation_summary(self._remember(presentation)),
            "already_open": False,
        }

    @action("save")
    def save(self, path: str | None = None) -> dict[str, Any]:
        """Saves the presentation (``Save``) or saves it as a new file (``SaveAs``)."""
        presentation = self.presentation()

        if path:
            target = self.resolve_target_path(path)
            with self.alerts_suppressed():
                presentation.SaveAs(
                    target,
                    save_format_for(target, PP_SAVE_FORMATS, PP_SAVE_FORMATS[".pptx"]),
                )
        elif not presentation.Path:
            raise InvalidReferenceError(
                "The presentation has no file yet - pass the path parameter"
            )
        else:
            presentation.Save()

        # SaveAs repoints the presentation at a new file - the pinned path has
        # to follow, otherwise presentation() will not find it again.
        return self._presentation_summary(self._remember(presentation))

    @action("close")
    def close(self, save: bool = True) -> dict[str, Any]:
        """Closes the presentation, optionally saving changes."""
        presentation = self.presentation()
        name = str(presentation.Name)

        if save:
            if not presentation.Path:
                raise InvalidReferenceError(
                    "The presentation was never saved - run ppt_save with a path first"
                )
            presentation.Save()
        else:
            presentation.Saved = MSO_TRUE

        presentation.Close()
        self._target_path = None
        return {"closed": name, "saved": bool(save)}

    @action("get_presentation_info")
    def get_presentation_info(self) -> dict[str, Any]:
        """Basic presentation metadata: slide size, theme, path."""
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
        """Slide list: index, title, layout and shape count."""
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
        """Full slide content: shapes, their positions, text and notes."""
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
        """Adds a slide with the given layout; ``index=None`` means at the end."""
        presentation = self.presentation()
        count = presentation.Slides.Count
        position = count + 1 if index is None else max(1, min(int(index), count + 1))
        layout_constant = lookup_constant(layout, PP_LAYOUTS, "layout")

        slide = presentation.Slides.Add(position, layout_constant)

        if title:
            title_shape = self._title_shape(slide)
            if title_shape is None:
                title_shape = slide.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, 60, 40, 600, 60)
            title_shape.TextFrame.TextRange.Text = _paragraph_text(title)

        self._goto_slide(position)
        return {
            "slide_index": position,
            "layout": layout,
            "slide_count": int(presentation.Slides.Count),
        }

    @action("delete_slide")
    def delete_slide(self, slide_index: int) -> dict[str, Any]:
        """Deletes the slide at the given index."""
        presentation = self.presentation()
        index = self.require_index(slide_index, presentation.Slides.Count, "slide_index")
        presentation.Slides(index).Delete()
        return {"deleted": index, "slide_count": int(presentation.Slides.Count)}

    @action("duplicate_slide")
    def duplicate_slide(self, slide_index: int) -> dict[str, Any]:
        """Duplicates a slide - the copy lands right after the original."""
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
        """Moves a slide to another position."""
        presentation = self.presentation()
        count = presentation.Slides.Count
        source = self.require_index(from_index, count, "from_index")
        target = self.require_index(to_index, count, "to_index")
        presentation.Slides(source).MoveTo(target)
        self._goto_slide(target)
        return {"from_index": source, "to_index": target}

    @action("set_title")
    def set_title(self, slide_index: int, text: str) -> dict[str, Any]:
        """Sets the slide title; if the layout has none, inserts a text box."""
        slide = self.slide(slide_index)
        shape = self._title_shape(slide)
        created = False

        if shape is None:
            shape = slide.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, 60, 40, 600, 60)
            shape.TextFrame.TextRange.Font.Size = 32
            shape.Name = TITLE_FALLBACK_NAME
            created = True

        shape.TextFrame.TextRange.Text = _paragraph_text(text)
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
        """Inserts a text box at the given spot on the slide (points)."""
        slide = self.slide(slide_index)
        shape = slide.Shapes.AddTextbox(
            MSO_TEXT_HORIZONTAL, float(left), float(top), float(width), float(height)
        )
        text_range = shape.TextFrame.TextRange
        text_range.Text = _paragraph_text(text)

        if font_size is not None:
            text_range.Font.Size = float(font_size)
        if bold:
            text_range.Font.Bold = MSO_TRUE
        if color is not None:
            text_range.Font.Color.RGB = parse_color(color)
        if align:
            text_range.ParagraphFormat.Alignment = PP_ALIGNMENTS.get(
                str(align).strip().lower(), PP_ALIGNMENTS["left"]
            )

        self._goto_slide(int(slide_index))
        return {"slide_index": int(slide_index), "shape_id": int(shape.Id)}

    @action("add_bullet_list")
    def add_bullet_list(
        self,
        slide_index: int,
        items: list[Any],
        placeholder: Any = "content",
    ) -> dict[str, Any]:
        """Fills a placeholder with a bulleted list, nesting supported.

        ``items`` accepts plain text (``"Point"``) or dictionaries with an
        wciecia (``{"text": "Podpunkt", "level": 2}``).
        """
        if not isinstance(items, list) or not items:
            raise InvalidReferenceError("List 'items' cannot be empty")

        slide = self.slide(slide_index)
        frame = self._placeholder_frame(slide, placeholder)

        entries = _normalize_outline(items)

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
        """Replaces text on one slide or across the whole presentation."""
        if not old_text:
            raise InvalidReferenceError("Parameter old_text cannot be empty")

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
        """Sets the speaker notes for a slide."""
        slide = self.slide(slide_index)
        try:
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = (
                _paragraph_text(text)
            )
        except com_error as exc:
            raise UnsupportedOperationError(
                "The slide has no room for speaker notes"
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
        """Formats all the text of the given shape."""
        slide = self.slide(slide_index)
        shape = self._find_shape(slide, shape_id)

        if not shape.HasTextFrame:
            raise InvalidReferenceError("The given shape contains no text")

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
        """Applies a theme from a ``.thmx``/``.potx`` file or the Office gallery."""
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
            f"Theme '{name_or_path}' not found. COM takes a path to a .thmx or "
            ".potx file - pass a full path or a theme name from the Office gallery."
        )

    @action("set_background")
    def set_background(
        self,
        slide_index: int,
        color: Any = None,
        image_path: str | None = None,
    ) -> dict[str, Any]:
        """Sets the slide background - a solid colour or an image."""
        if color is None and not image_path:
            raise InvalidReferenceError("Pass a colour or a path to a background image")

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
        """Changes the slide layout - by master layout name or a standard name."""
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
        """Inserts an image; omitting width/height keeps the original proportions."""
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
        """Inserts a chart and fills its worksheet with data.

        ``series_data`` accepts a dictionary ``{"name": [values]}``, a list of
        slownikow ``{"name": ..., "values": [...]}`` albo sama liste serii.
        """
        if not categories:
            raise InvalidReferenceError("List 'categories' cannot be empty")

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
        """Writes data into the chart's embedded sheet and sets the source range."""
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
        chart.SetSourceData(f"='{worksheet.Name}'!{com_address(data_range)}")

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
        """Inserts a table and fills it with data (extra cells are skipped)."""
        rows, cols = int(rows), int(cols)
        if rows < 1 or cols < 1:
            raise InvalidReferenceError("A table needs at least 1 row and 1 column")

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
                cell.Shape.TextFrame.TextRange.Text = _paragraph_text(value)
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
        line_color: Any = None,
        line_width: float | None = None,
    ) -> dict[str, Any]:
        """Inserts a shape (rectangle, arrow, star...) with optional text.

        ``line_color="none"`` removes the outline - without it the shape gets a
        default theme border, which rarely matches your own colours.
        """
        slide = self.slide(slide_index)
        shape_constant = lookup_constant(shape_type, SHAPE_TYPES, "shape_type")
        shape = slide.Shapes.AddShape(
            shape_constant, float(left), float(top), float(width), float(height)
        )

        if fill_color is not None:
            if str(fill_color).strip().lower() == "none":
                shape.Fill.Visible = MSO_FALSE
            else:
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = parse_color(fill_color)
        if line_color is not None:
            if str(line_color).strip().lower() == "none":
                shape.Line.Visible = MSO_FALSE
            else:
                shape.Line.Visible = MSO_TRUE
                shape.Line.ForeColor.RGB = parse_color(line_color)
        if line_width is not None:
            shape.Line.Visible = MSO_TRUE
            shape.Line.Weight = float(line_width)
        if text:
            shape.TextFrame.TextRange.Text = _paragraph_text(text)

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "shape_type": shape_type,
        }


    @action("delete_shape")
    def delete_shape(self, slide_index: int, shape_id: Any) -> dict[str, Any]:
        """Deletes a shape from the slide - by id or by name."""
        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)
        removed = {"shape_id": int(shape.Id), "shape_name": to_python(shape.Name)}
        shape.Delete()

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            **removed,
            "shapes_left": int(slide.Shapes.Count),
        }

    @action("set_shape_position")
    def set_shape_position(
        self,
        slide_index: int,
        shape_id: Any,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        rotation: float | None = None,
    ) -> dict[str, Any]:
        """Moves, scales and rotates an existing shape (given fields, in points)."""
        if all(value is None for value in (left, top, width, height, rotation)):
            raise InvalidReferenceError(
                "Podaj przynajmniej jedno z: left, top, width, height, rotation"
            )

        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)

        # With the aspect ratio locked, setting the width also changes the
        # height - we lift the lock when both dimensions are given at once.
        previous_lock: Any = None
        if width is not None and height is not None:
            try:
                previous_lock = shape.LockAspectRatio
                shape.LockAspectRatio = MSO_FALSE
            except com_error:
                previous_lock = None

        try:
            if left is not None:
                shape.Left = float(left)
            if top is not None:
                shape.Top = float(top)
            if width is not None:
                shape.Width = float(width)
            if height is not None:
                shape.Height = float(height)
            if rotation is not None:
                shape.Rotation = float(rotation)
        finally:
            if previous_lock is not None:
                try:
                    shape.LockAspectRatio = previous_lock
                except com_error:
                    pass

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "left": round(float(shape.Left), 2),
            "top": round(float(shape.Top), 2),
            "width": round(float(shape.Width), 2),
            "height": round(float(shape.Height), 2),
            "rotation": round(float(shape.Rotation), 2),
        }

    @action("set_shape_order")
    def set_shape_order(
        self, slide_index: int, shape_id: Any, order: str = "front"
    ) -> dict[str, Any]:
        """Changes the shape layer: front, back, forward or backward."""
        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)
        shape.ZOrder(lookup_constant(order, MSO_ZORDER, "order"))

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "order": order,
            "z_order_position": int(shape.ZOrderPosition),
            "shapes_on_slide": int(slide.Shapes.Count),
        }

    @action("export_slide")
    def export_slide(
        self,
        slide_index: int,
        path: str,
        width: int | None = None,
        height: int | None = None,
    ) -> dict[str, Any]:
        """Saves the slide as an image (PNG/JPG/GIF/BMP/WMF/EMF by extension).

        Without ``width``/``height`` the image is 1920 px wide and the height is
        computed from the slide proportions.
        """
        slide = self.slide(slide_index)
        target = self.resolve_target_path(path)

        extension = os.path.splitext(target)[1].lower()
        if extension not in PP_EXPORT_FILTERS:
            raise InvalidReferenceError(
                f"Unsupported image extension: {extension or '(none)'}. "
                f"Dostepne: {', '.join(sorted(PP_EXPORT_FILTERS))}"
            )

        setup = self.presentation().PageSetup
        ratio = float(setup.SlideHeight) / float(setup.SlideWidth)
        if width is None and height is None:
            width = 1920
            height = int(round(width * ratio))
        elif height is None:
            height = int(round(int(width) * ratio))
        elif width is None:
            width = int(round(int(height) / ratio))

        slide.Export(target, PP_EXPORT_FILTERS[extension], int(width), int(height))

        return {
            "slide_index": int(slide_index),
            "path": target,
            "format": PP_EXPORT_FILTERS[extension],
            "width": int(width),
            "height": int(height),
            "size_bytes": os.path.getsize(target) if os.path.isfile(target) else None,
        }

    @action("export_pdf")
    def export_pdf(self, path: str, embed_fonts: bool = True) -> dict[str, Any]:
        """Exports the whole presentation to PDF without changing the current file.

        Uses ``SaveCopyAs``, not ``ExportAsFixedFormat`` - the latter cannot be
        called through pywin32 (the wrapper puts ``PyOleEmpty`` in the
        ``ExternalExporter`` parameter, which cannot be converted to COM).
        ``SaveAs`` is out too, because it would repoint the presentation open in
        PowerPoint at the PDF file.
        """
        presentation = self.presentation()
        target = self.resolve_target_path(path)

        with self.alerts_suppressed():
            presentation.SaveCopyAs(
                target,
                PP_SAVE_FORMATS[".pdf"],
                MSO_TRUE if embed_fonts else MSO_FALSE,
            )

        return {
            "path": target,
            "slide_count": int(presentation.Slides.Count),
            "embed_fonts": bool(embed_fonts),
            "size_bytes": os.path.getsize(target) if os.path.isfile(target) else None,
        }

    def _shape_range(self, slide: Any, shape_ids: Any, minimum: int = 2) -> Any:
        """``ShapeRange`` from a list of ids/names - the basis of grouping and aligning."""
        if not isinstance(shape_ids, (list, tuple)) or len(shape_ids) < minimum:
            raise InvalidReferenceError(
                f"Podaj liste co najmniej {minimum} ksztaltow (id albo nazw)"
            )

        indexes: list[int] = []
        for shape_id in shape_ids:
            wanted = self._resolve_shape(slide, shape_id)
            for index in range(1, slide.Shapes.Count + 1):
                if int(slide.Shapes(index).Id) == int(wanted.Id):
                    indexes.append(index)
                    break

        if len(indexes) < minimum:
            raise InvalidReferenceError(
                f"Found only {len(indexes)} of the {len(shape_ids)} shapes given"
            )
        return slide.Shapes.Range(indexes)

    @action("group_shapes")
    def group_shapes(
        self, slide_index: int, shape_ids: list[Any], name: str | None = None
    ) -> dict[str, Any]:
        """Groups shapes - from now on they move and animate as one."""
        slide = self.slide(slide_index)
        group = self._shape_range(slide, shape_ids).Group()

        if name:
            group.Name = str(name)

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(group.Id),
            "name": to_python(group.Name),
            "grouped": len(shape_ids),
        }

    @action("ungroup_shapes")
    def ungroup_shapes(self, slide_index: int, shape_id: Any) -> dict[str, Any]:
        """Breaks a group back into individual shapes."""
        slide = self.slide(slide_index)
        group = self._resolve_shape(slide, shape_id)

        try:
            parts = group.Ungroup()
        except com_error as exc:
            raise InvalidReferenceError(
                f"Shape {shape_id!r} is not a group"
            ) from exc

        return {
            "slide_index": int(slide_index),
            "shapes": [int(parts(i).Id) for i in range(1, int(parts.Count) + 1)],
            "count": int(parts.Count),
        }

    @action("align_shapes")
    def align_shapes(
        self,
        slide_index: int,
        shape_ids: list[Any],
        align: str,
        relative_to_slide: bool = False,
    ) -> dict[str, Any]:
        """Aligns shapes: left, center, right, top, middle, bottom.

        ``relative_to_slide=True`` aligns to the slide edges rather than to
        wzgledem siebie nawzajem.
        """
        slide = self.slide(slide_index)
        shape_range = self._shape_range(
            slide, shape_ids, minimum=1 if relative_to_slide else 2
        )
        shape_range.Align(
            lookup_constant(align, MSO_ALIGN, "align"),
            MSO_TRUE if relative_to_slide else MSO_FALSE,
        )

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "align": align,
            "shapes": len(shape_ids),
            "relative_to_slide": bool(relative_to_slide),
        }

    @action("distribute_shapes")
    def distribute_shapes(
        self,
        slide_index: int,
        shape_ids: list[Any],
        direction: str = "horizontal",
        relative_to_slide: bool = False,
    ) -> dict[str, Any]:
        """Spreads shapes at equal intervals, horizontally or vertically."""
        slide = self.slide(slide_index)
        shape_range = self._shape_range(slide, shape_ids, minimum=3)
        shape_range.Distribute(
            lookup_constant(direction, MSO_DISTRIBUTE, "direction"),
            MSO_TRUE if relative_to_slide else MSO_FALSE,
        )

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "direction": direction,
            "shapes": len(shape_ids),
            "relative_to_slide": bool(relative_to_slide),
        }

    @action("add_hyperlink")
    def add_hyperlink(
        self,
        slide_index: int,
        shape_id: Any,
        url: str | None = None,
        target_slide: int | None = None,
        tooltip: str | None = None,
    ) -> dict[str, Any]:
        """Attaches a link to a shape - external (``url``) or to a slide.

        Passing only ``tooltip`` without ``url``/``target_slide`` makes no sense,
        because the hint only shows on an active link.
        """
        if url and target_slide is not None:
            raise InvalidReferenceError("Pass 'url' or 'target_slide', not both")
        if not url and target_slide is None:
            raise InvalidReferenceError("Podaj 'url' albo 'target_slide'")

        presentation = self.presentation()
        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)

        settings = shape.ActionSettings(PP_MOUSE_CLICK)
        settings.Action = PP_ACTION_HYPERLINK
        hyperlink = settings.Hyperlink

        if url:
            hyperlink.Address = str(url)
            hyperlink.SubAddress = ""
            applied: dict[str, Any] = {"url": str(url)}
        else:
            index = self.require_index(
                target_slide, presentation.Slides.Count, "target_slide"
            )
            target = presentation.Slides(index)
            title_shape = self._title_shape(target)
            title = ""
            if title_shape is not None:
                try:
                    title = str(title_shape.TextFrame.TextRange.Text)
                except com_error:
                    title = ""
            hyperlink.Address = ""
            hyperlink.SubAddress = f"{int(target.SlideID)},{index},{title}"
            applied = {"target_slide": index, "target_title": title or None}

        if tooltip:
            hyperlink.ScreenTip = str(tooltip)
            applied["tooltip"] = str(tooltip)

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            **applied,
        }

    @action("set_headers_footers")
    def set_headers_footers(
        self,
        slide_index: int | None = None,
        footer_text: str | None = None,
        show_footer: bool | None = None,
        show_slide_number: bool | None = None,
        show_date: bool | None = None,
    ) -> dict[str, Any]:
        """Footer, slide number and date; without ``slide_index`` on every slide.

        Passing only ``footer_text`` turns the footer on - otherwise the text
        would stay invisible until the field is ticked by hand in PowerPoint.
        """
        if all(
            value is None
            for value in (footer_text, show_footer, show_slide_number, show_date)
        ):
            raise InvalidReferenceError("No field given to change")

        presentation = self.presentation()
        if slide_index is None:
            indexes = list(range(1, presentation.Slides.Count + 1))
        else:
            indexes = [
                self.require_index(slide_index, presentation.Slides.Count, "slide_index")
            ]

        text_on_master = False
        for index in indexes:
            headers = presentation.Slides(index).HeadersFooters
            if footer_text is not None:
                try:
                    headers.Footer.Text = _paragraph_text(footer_text)
                except com_error:
                    # A layout without a footer placeholder (e.g. "blank")
                    # rejects the text on the slide - it then lives on the master.
                    presentation.SlideMaster.HeadersFooters.Footer.Text = (
                        _paragraph_text(footer_text)
                    )
                    text_on_master = True
                if show_footer is None:
                    headers.Footer.Visible = MSO_TRUE
            if show_footer is not None:
                headers.Footer.Visible = MSO_TRUE if show_footer else MSO_FALSE
            if show_slide_number is not None:
                headers.SlideNumber.Visible = MSO_TRUE if show_slide_number else MSO_FALSE
            if show_date is not None:
                headers.DateAndTime.Visible = MSO_TRUE if show_date else MSO_FALSE

        return {
            "slides": indexes,
            "footer_text": footer_text,
            "show_footer": show_footer,
            "show_slide_number": show_slide_number,
            "show_date": show_date,
            "text_on_master": text_on_master,
        }

    def _theme(self) -> Any:
        """Slide master theme - the single place the whole palette comes from."""
        return self.presentation().SlideMaster.Theme

    @action("get_theme")
    def get_theme(self) -> dict[str, Any]:
        """Returns the theme palette and fonts - to check what is set."""
        theme = self._theme()
        scheme = theme.ThemeColorScheme
        fonts = theme.ThemeFontScheme

        return {
            "colors": {
                name: bgr_to_hex(scheme.Colors(index).RGB)
                for name, index in MSO_THEME_COLORS.items()
                if index <= int(scheme.Count)
            },
            "major_font": to_python(fonts.MajorFont(MSO_THEME_LATIN).Name),
            "minor_font": to_python(fonts.MinorFont(MSO_THEME_LATIN).Name),
            "theme_name": to_python(self.presentation().SlideMaster.Design.Name),
        }

    @action("set_theme_colors")
    def set_theme_colors(self, colors: dict[str, Any]) -> dict[str, Any]:
        """Podmienia kolory w palecie motywu.

        ``colors`` to slownik ``{"accent1": "#10A37F", "dark1": "#0B1014"}``.
        Names come from ``MSO_THEME_COLORS``: ``dark1``/``text1``,
        ``light1``/``background1``, ``dark2``, ``light2``, ``accent1``-``accent6``,
        ``hyperlink``, ``followed_hyperlink``.
        """
        if not isinstance(colors, dict) or not colors:
            raise InvalidReferenceError(
                "Podaj slownik kolorow, np. {\"accent1\": \"#10A37F\"}"
            )

        scheme = self._theme().ThemeColorScheme
        applied: dict[str, Any] = {}
        for name, value in colors.items():
            index = lookup_constant(name, MSO_THEME_COLORS, "theme colour name")
            scheme.Colors(index).RGB = parse_color(value)
            applied[str(name)] = bgr_to_hex(scheme.Colors(index).RGB)

        return {"colors": applied}

    @action("set_theme_fonts")
    def set_theme_fonts(
        self, major: str | None = None, minor: str | None = None
    ) -> dict[str, Any]:
        """Sets theme fonts: ``major`` for headings, ``minor`` for body text."""
        if not major and not minor:
            raise InvalidReferenceError("Podaj 'major', 'minor' albo oba")

        fonts = self._theme().ThemeFontScheme
        if major:
            fonts.MajorFont(MSO_THEME_LATIN).Name = str(major)
        if minor:
            fonts.MinorFont(MSO_THEME_LATIN).Name = str(minor)

        return {
            "major_font": to_python(fonts.MajorFont(MSO_THEME_LATIN).Name),
            "minor_font": to_python(fonts.MinorFont(MSO_THEME_LATIN).Name),
        }

    @action("set_master_background")
    def set_master_background(
        self,
        color: Any = None,
        image_path: str | None = None,
        apply_to_slides: bool = True,
    ) -> dict[str, Any]:
        """Sets the background on the slide master - once for the whole deck.

        ``apply_to_slides=True`` turns on ``FollowMasterBackground`` for slides,
        so those with their own background from ``set_background`` return to it.
        """
        if color is None and not image_path:
            raise InvalidReferenceError("Pass a colour or a path to a background image")

        presentation = self.presentation()
        fill = presentation.SlideMaster.Background.Fill

        if image_path:
            resolved = self.resolve_existing_path(image_path)
            fill.UserPicture(resolved)
            applied: dict[str, Any] = {"image_path": resolved}
        else:
            fill.Solid()
            fill.ForeColor.RGB = parse_color(color)
            applied = {"color": bgr_to_hex(fill.ForeColor.RGB)}

        followed = 0
        if apply_to_slides:
            for index in range(1, presentation.Slides.Count + 1):
                presentation.Slides(index).FollowMasterBackground = MSO_TRUE
                followed += 1

        return {**applied, "slides_following_master": followed}

    @action("set_shape_format")
    def set_shape_format(
        self,
        slide_index: int,
        shape_id: Any,
        fill_color: Any = None,
        fill_transparency: float | None = None,
        gradient_from: Any = None,
        gradient_to: Any = None,
        gradient_style: str = "vertical",
        line_color: Any = None,
        line_width: float | None = None,
        line_dash: str | None = None,
        shadow: bool | None = None,
        shadow_color: Any = None,
        shadow_blur: float | None = None,
        shadow_offset_x: float | None = None,
        shadow_offset_y: float | None = None,
        shadow_transparency: float | None = None,
        corner_radius: float | None = None,
    ) -> dict[str, Any]:
        """Look of an existing shape: gradient, transparency, shadow, outline.

        ``gradient_from`` + ``gradient_to`` wlaczaja gradient dwukolorowy
        (``gradient_style``: horizontal, vertical, diagonal_up, diagonal_down,
        from_corner, from_center). ``corner_radius`` 0.0-0.5 dziala na ksztaltach
        z uchwytem regulacji, np. ``rounded_rectangle``.
        """
        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)
        applied: dict[str, Any] = {}

        if gradient_from is not None and gradient_to is not None:
            style = lookup_constant(gradient_style, MSO_GRADIENT_STYLES, "gradient_style")
            shape.Fill.TwoColorGradient(style, 1)
            shape.Fill.ForeColor.RGB = parse_color(gradient_from)
            shape.Fill.BackColor.RGB = parse_color(gradient_to)
            applied["gradient"] = {
                "from": bgr_to_hex(shape.Fill.ForeColor.RGB),
                "to": bgr_to_hex(shape.Fill.BackColor.RGB),
                "style": gradient_style,
            }
        elif gradient_from is not None or gradient_to is not None:
            raise InvalidReferenceError(
                "A gradient needs both colours: gradient_from and gradient_to"
            )
        elif fill_color is not None:
            if str(fill_color).strip().lower() == "none":
                shape.Fill.Visible = MSO_FALSE
                applied["fill"] = "none"
            else:
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = parse_color(fill_color)
                applied["fill"] = bgr_to_hex(shape.Fill.ForeColor.RGB)

        if fill_transparency is not None:
            shape.Fill.Transparency = _unit_fraction(fill_transparency, "fill_transparency")
            applied["fill_transparency"] = round(float(shape.Fill.Transparency), 3)

        if line_color is not None:
            if str(line_color).strip().lower() == "none":
                shape.Line.Visible = MSO_FALSE
                applied["line"] = "none"
            else:
                shape.Line.Visible = MSO_TRUE
                shape.Line.ForeColor.RGB = parse_color(line_color)
                applied["line"] = bgr_to_hex(shape.Line.ForeColor.RGB)
        if line_width is not None:
            shape.Line.Visible = MSO_TRUE
            shape.Line.Weight = float(line_width)
            applied["line_width"] = float(line_width)
        if line_dash is not None:
            shape.Line.Visible = MSO_TRUE
            shape.Line.DashStyle = lookup_constant(line_dash, MSO_LINE_DASHES, "line_dash")
            applied["line_dash"] = line_dash

        shadow_requested = any(
            value is not None
            for value in (
                shadow_color,
                shadow_blur,
                shadow_offset_x,
                shadow_offset_y,
                shadow_transparency,
            )
        )
        if shadow is not None or shadow_requested:
            visible = MSO_TRUE if (shadow or (shadow is None and shadow_requested)) else MSO_FALSE
            shape.Shadow.Visible = visible
            if visible == MSO_TRUE:
                shape.Shadow.Style = MSO_SHADOW_OUTER
                if shadow_color is not None:
                    shape.Shadow.ForeColor.RGB = parse_color(shadow_color)
                if shadow_blur is not None:
                    shape.Shadow.Blur = float(shadow_blur)
                if shadow_offset_x is not None:
                    shape.Shadow.OffsetX = float(shadow_offset_x)
                if shadow_offset_y is not None:
                    shape.Shadow.OffsetY = float(shadow_offset_y)
                if shadow_transparency is not None:
                    shape.Shadow.Transparency = _unit_fraction(
                        shadow_transparency, "shadow_transparency"
                    )
            applied["shadow"] = visible == MSO_TRUE

        if corner_radius is not None:
            if int(shape.Adjustments.Count) < 1:
                raise InvalidReferenceError(
                    "This shape has no adjustment handle - corner_radius works "
                    "np. na rounded_rectangle"
                )
            # Adjustments to parametryzowana wlasciwosc COM; pywin32 wystawia ja
            # writable through SetItem; plain assignment does not work.
            shape.Adjustments.SetItem(1, _unit_fraction(corner_radius, "corner_radius"))
            applied["corner_radius"] = round(float(shape.Adjustments.Item(1)), 4)

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "applied": applied,
        }

    @action("set_paragraph_format")
    def set_paragraph_format(
        self,
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
        """Paragraph typography: line spacing, gaps, alignment, vertical anchor.

        ``paragraph`` left empty covers all the text of the shape; with a number
        (1-based) only that paragraph. ``line_spacing`` is a multiple of the line
        height (1.0 = single), gaps are in points.
        """
        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)

        if not shape.HasTextFrame:
            raise InvalidReferenceError("This shape has no text frame")

        frame = shape.TextFrame
        text_range = frame.TextRange
        if paragraph is not None:
            count = int(text_range.Paragraphs().Count)
            index = self.require_index(paragraph, count, "paragraph")
            text_range = text_range.Paragraphs(index)

        applied: dict[str, Any] = {}
        paragraph_format = text_range.ParagraphFormat

        if line_spacing is not None:
            paragraph_format.SpaceWithin = float(line_spacing)
            applied["line_spacing"] = float(line_spacing)
        if space_before is not None:
            paragraph_format.SpaceBefore = float(space_before)
            applied["space_before"] = float(space_before)
        if space_after is not None:
            paragraph_format.SpaceAfter = float(space_after)
            applied["space_after"] = float(space_after)
        if alignment is not None:
            paragraph_format.Alignment = lookup_constant(
                alignment, PP_ALIGNMENTS, "alignment"
            )
            applied["alignment"] = alignment

        if vertical_anchor is not None:
            frame.VerticalAnchor = lookup_constant(
                vertical_anchor, MSO_ANCHORS, "vertical_anchor"
            )
            applied["vertical_anchor"] = vertical_anchor
        if autosize is not None:
            frame.AutoSize = PP_AUTOSIZE_FIT if autosize else PP_AUTOSIZE_NONE
            applied["autosize"] = bool(autosize)
        if word_wrap is not None:
            frame.WordWrap = MSO_TRUE if word_wrap else MSO_FALSE
            applied["word_wrap"] = bool(word_wrap)
        if margin is not None:
            for side in ("MarginLeft", "MarginRight", "MarginTop", "MarginBottom"):
                setattr(frame, side, float(margin))
            applied["margin"] = float(margin)

        if not applied:
            raise InvalidReferenceError("No field given to change")

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "paragraph": int(paragraph) if paragraph is not None else "all",
            "applied": applied,
        }

    @action("format_chart")
    def format_chart(
        self,
        slide_index: int,
        shape_id: Any,
        series_colors: list[Any] | None = None,
        text_color: Any = None,
        background: Any = None,
        legend: Any = None,
        data_labels: bool | None = None,
        gridlines: bool | None = None,
        title: str | None = None,
        value_axis_min: float | None = None,
        value_axis_max: float | None = None,
    ) -> dict[str, Any]:
        """Tunes a chart to the slide colours.

        ``background="none"`` makes the chart background transparent, ``legend``
        ``False`` albo pozycje (``bottom``, ``top``, ``left``, ``right``),
        ``text_color`` paints axes, legend and title at once.
        """
        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)

        if not shape.HasChart:
            raise InvalidReferenceError(
                f"Shape {shape_id!r} is not a chart - use ppt_add_chart"
            )

        applied = apply_chart_format(
            shape.Chart,
            series_colors=series_colors,
            text_color=text_color,
            background=background,
            legend=legend,
            data_labels=data_labels,
            gridlines=gridlines,
            title=title,
            value_axis_min=value_axis_min,
            value_axis_max=value_axis_max,
        )

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "applied": applied,
        }

    @action("add_media")
    def add_media(
        self,
        slide_index: int,
        media_path: str,
        left: float,
        top: float,
        width: float | None = None,
        height: float | None = None,
        autoplay: bool = False,
    ) -> dict[str, Any]:
        """Inserts video or audio as an object embedded in the presentation.

        ``autoplay=True`` attaches a play effect starting with the previous one,
        instead of waiting for a click.
        """
        slide = self.slide(slide_index)
        target = self.resolve_existing_path(media_path)

        try:
            shape = slide.Shapes.AddMediaObject2(
                FileName=target,
                LinkToFile=MSO_FALSE,
                SaveWithDocument=MSO_TRUE,
                Left=float(left),
                Top=float(top),
                Width=float(width) if width is not None else -1,
                Height=float(height) if height is not None else -1,
            )
        except (com_error, AttributeError):
            shape = slide.Shapes.AddMediaObject(
                target,
                float(left),
                float(top),
                float(width) if width is not None else -1,
                float(height) if height is not None else -1,
            )

        media_type = None
        try:
            media_type = {2: "sound", 3: "movie"}.get(int(shape.MediaType), "other")
        except com_error:
            pass

        if autoplay:
            try:
                slide.TimeLine.MainSequence.AddEffect(
                    Shape=shape,
                    effectId=MSO_ANIM_MEDIA_PLAY,
                    Level=0,
                    trigger=MSO_ANIM_TRIGGERS["with_previous"],
                )
            except com_error:
                pass

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "media_type": media_type,
            "path": target,
            "width": round(float(shape.Width), 2),
            "height": round(float(shape.Height), 2),
            "autoplay": bool(autoplay),
        }

    @action("list_smartart_layouts")
    def list_smartart_layouts(
        self, search: str | None = None, category: str | None = None
    ) -> dict[str, Any]:
        """Available SmartArt layouts: key, name and category.

        ``name`` is localised - a non-English Office returns translated layout
        names - so pick a layout by ``key`` instead: the tail of the URN
        identifier, which is the same in every language version.
        Kategorie tez sa niezalezne od jezyka: list, process, cycle, hierarchy,
        relationship, matrix, pyramid, picture.
        """
        layouts = self.app.SmartArtLayouts
        needle = str(search).strip().lower() if search else None
        wanted_category = str(category).strip().lower() if category else None

        found: list[dict[str, Any]] = []
        for index in range(1, int(layouts.Count) + 1):
            entry = self._smartart_entry(layouts(index), index)
            if wanted_category and entry["category"] != wanted_category:
                continue
            if needle and needle not in entry["key"].lower() and (
                needle not in str(entry["name"]).lower()
            ):
                continue
            found.append(entry)

        return {"count": len(found), "total": int(layouts.Count), "layouts": found}

    @staticmethod
    def _smartart_entry(layout: Any, index: int) -> dict[str, Any]:
        identifier = str(to_python(layout.Id) or "")
        return {
            "index": index,
            "key": identifier.rsplit("/", 1)[-1] if identifier else "",
            "name": to_python(layout.Name),
            "category": str(to_python(getattr(layout, "Category", "")) or "").lower(),
        }

    def _smartart_layout(self, layout: Any) -> Any:
        """SmartArt layout by number, URN key or name (localised names included)."""
        layouts = self.app.SmartArtLayouts
        if isinstance(layout, int) or str(layout).strip().isdigit():
            index = self.require_index(layout, int(layouts.Count), "layout")
            return layouts(index)

        needle = str(layout).strip().lower()
        entries = [
            (self._smartart_entry(layouts(index), index), layouts(index))
            for index in range(1, int(layouts.Count) + 1)
        ]

        for match in (
            lambda entry: entry["key"].lower() == needle,
            lambda entry: str(entry["name"]).strip().lower() == needle,
            lambda entry: needle in entry["key"].lower(),
            lambda entry: needle in str(entry["name"]).lower(),
        ):
            for entry, com_layout in entries:
                if match(entry):
                    return com_layout

        raise InvalidReferenceError(
            f"No SmartArt layout matches {layout!r}. Names are localised - use "
            "ppt_list_smartart_layouts to see the keys "
            "(np. 'bProcess3', 'hierarchy1') albo filtruj po kategorii."
        )

    @action("add_smartart")
    def add_smartart(
        self,
        slide_index: int,
        layout: Any,
        items: list[Any],
        left: float,
        top: float,
        width: float,
        height: float,
    ) -> dict[str, Any]:
        """Inserts a SmartArt diagram and fills it with text.

        ``items`` przyjmuje liste tekstow albo slownikow
        ``{"text": ..., "level": 1}`` - level 2 and deeper creates child nodes.
        """
        slide = self.slide(slide_index)
        chosen = self._smartart_layout(layout)
        shape = slide.Shapes.AddSmartArt(
            chosen, float(left), float(top), float(width), float(height)
        )

        entries = _normalize_outline(items)
        smart_art = shape.SmartArt

        while int(smart_art.AllNodes.Count) > 0:
            smart_art.AllNodes.Item(1).Delete()

        # Child nodes come from AddNode(below) on the parent. Demote() on a node
        # z AllNodes.Add() czesc ukladow (np. hierarchy1) odrzuca komunikatem
        # "operation not supported by the current object".
        last_at_level: dict[int, Any] = {}
        added = 0
        for text, level in entries:
            parent = last_at_level.get(level - 1) if level > 1 else None
            if parent is None:
                node = smart_art.AllNodes.Add()
                level = 1
            else:
                node = parent.AddNode(MSO_SMARTART_NODE_BELOW)

            node.TextFrame2.TextRange.Text = _paragraph_text(text)
            last_at_level[level] = node
            for deeper in [key for key in last_at_level if key > level]:
                del last_at_level[deeper]
            added += 1

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "layout": to_python(chosen.Name),
            "nodes": added,
        }

    @action("list_sections")
    def list_sections(self) -> dict[str, Any]:
        """Presentation sections with their slide ranges."""
        properties = self.presentation().SectionProperties
        count = int(properties.Count)

        return {
            "count": count,
            "sections": [
                {
                    "index": index,
                    "name": to_python(properties.Name(index)),
                    "first_slide": int(properties.FirstSlide(index)),
                    "slides": int(properties.SlidesCount(index)),
                }
                for index in range(1, count + 1)
            ],
        }

    @action("add_section")
    def add_section(self, name: str, before_slide: int = 1) -> dict[str, Any]:
        """Creates a section starting at the given slide."""
        presentation = self.presentation()
        index = self.require_index(
            before_slide, presentation.Slides.Count, "before_slide"
        )
        section_index = presentation.SectionProperties.AddBeforeSlide(index, str(name))

        return {
            "section_index": int(section_index),
            "name": str(name),
            "first_slide": index,
        }

    @action("delete_section")
    def delete_section(
        self, section_index: int, delete_slides: bool = False
    ) -> dict[str, Any]:
        """Deletes a section; ``delete_slides=True`` removes its slides too."""
        presentation = self.presentation()
        properties = presentation.SectionProperties
        index = self.require_index(section_index, int(properties.Count), "section_index")
        name = to_python(properties.Name(index))

        properties.Delete(index, MSO_TRUE if delete_slides else MSO_FALSE)

        return {
            "deleted": index,
            "name": name,
            "slides_deleted": bool(delete_slides),
            "sections_left": int(properties.Count),
            "slide_count": int(presentation.Slides.Count),
        }

    @action("slideshow")
    def slideshow(
        self, command: str = "start", slide_index: int | None = None
    ) -> dict[str, Any]:
        """Steruje pokazem: ``start``, ``stop`` albo ``goto`` (z ``slide_index``)."""
        presentation = self.presentation()
        wanted = str(command).strip().lower()

        if wanted == "start":
            settings = presentation.SlideShowSettings
            if slide_index is not None:
                index = self.require_index(
                    slide_index, presentation.Slides.Count, "slide_index"
                )
                settings.RangeType = 2  # ppShowSlideRange
                settings.StartingSlide = index
                settings.EndingSlide = int(presentation.Slides.Count)
            settings.Run()
            return {"command": "start", "running": True}

        if wanted in ("stop", "exit", "end"):
            try:
                presentation.SlideShowWindow.View.Exit()
            except com_error as exc:
                raise InvalidReferenceError("The slide show is not running") from exc
            return {"command": "stop", "running": False}

        if wanted == "goto":
            if slide_index is None:
                raise InvalidReferenceError("'goto' requires the slide_index parameter")
            index = self.require_index(
                slide_index, presentation.Slides.Count, "slide_index"
            )
            # Right after Run() PowerPoint is still building the show window and
            # wywolania (RPC_E_CALL_REJECTED) - jedno ponowienie wystarcza.
            for attempt in range(2):
                try:
                    presentation.SlideShowWindow.View.GotoSlide(index)
                    break
                except com_error as exc:
                    if attempt:
                        raise InvalidReferenceError(
                            "The slide show is not running, or PowerPoint has "
                            "not opened it yet"
                        ) from exc
                    time.sleep(0.6)
            return {"command": "goto", "slide_index": index, "running": True}

        raise InvalidReferenceError(
            f"Unknown command: {command!r}. Available: start, stop, goto"
        )

    @action("copy_slide_to")
    def copy_slide_to(
        self, slide_index: int, target_path: str, position: int | None = None
    ) -> dict[str, Any]:
        """Copies a slide into another presentation (an existing .pptx file).

        Uses ``Slides.InsertFromFile`` rather than the clipboard - the clipboard
        can be busy with the user's own copy and spoil the result unpredictably.
        """
        presentation = self.presentation()
        index = self.require_index(slide_index, presentation.Slides.Count, "slide_index")
        target = self.resolve_existing_path(target_path)

        if not presentation.Path:
            raise DocumentNotFoundError(
                "The source presentation has no file yet - save it with ppt_save"
            )
        source = str(presentation.FullName)
        if os.path.normcase(source) == os.path.normcase(target):
            raise InvalidReferenceError(
                "Source and target are the same file - use ppt_duplicate_slide"
            )

        with self.alerts_suppressed():
            presentation.Save()

        app = self.app
        opened = None
        for i in range(1, app.Presentations.Count + 1):
            if os.path.normcase(str(app.Presentations(i).FullName)) == os.path.normcase(target):
                opened = app.Presentations(i)
                break
        destination = opened or app.Presentations.Open(
            target, ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE
        )

        where = int(destination.Slides.Count) if position is None else max(
            0, self.require_index(position, int(destination.Slides.Count) + 1, "position") - 1
        )
        inserted = destination.Slides.InsertFromFile(source, where, index, index)

        with self.alerts_suppressed():
            destination.Save()

        return {
            "source_slide": index,
            "target_path": target,
            "inserted_at": int(where) + 1,
            "target_slide_count": int(destination.Slides.Count),
            "slides_inserted": int(inserted),
        }

    def _resolve_shape(self, slide: Any, shape_id: Any) -> Any:
        """Shape by id/name, but also understands ``title`` and ``content``."""
        if isinstance(shape_id, str):
            wanted = shape_id.strip().lower()
            if wanted in ("title",):
                shape = self._title_shape(slide)
                if shape is None:
                    raise InvalidReferenceError("The slide has no title placeholder")
                return shape
            if wanted in ("content", "body"):
                for index in range(1, slide.Shapes.Placeholders.Count + 1):
                    shape = slide.Shapes.Placeholders(index)
                    try:
                        placeholder_type = int(shape.PlaceholderFormat.Type)
                    except com_error:
                        continue
                    if placeholder_type in CONTENT_PLACEHOLDERS:
                        return shape
                raise InvalidReferenceError("The slide has no content placeholder")
        return self._find_shape(slide, shape_id)

    @action("add_animation")
    def add_animation(
        self,
        slide_index: int,
        shape_id: Any,
        effect: str = "fade",
        trigger: str = "after_previous",
        level: str = "shape",
        duration: float | None = None,
        delay: float | None = None,
        exit_effect: bool = False,
    ) -> dict[str, Any]:
        """Adds a shape animation to the slide's main sequence.

        ``shape_id`` is a shape id, its name, or the ``title`` / ``content``
        shortcut. ``level`` decides whether the whole shape animates
        (``shape``) or its text paragraph by paragraph (``by_paragraph``).
        ``exit_effect=True`` zamienia efekt wejscia na wyjscie.
        """
        slide = self.slide(slide_index)
        shape = self._resolve_shape(slide, shape_id)

        effect_constant = lookup_constant(effect, MSO_ANIM_EFFECTS, "effect")
        trigger_constant = lookup_constant(trigger, MSO_ANIM_TRIGGERS, "trigger")
        level_constant = lookup_constant(level, MSO_ANIM_LEVELS, "level")

        try:
            sequence = slide.TimeLine.MainSequence
        except (com_error, AttributeError) as exc:
            raise UnsupportedOperationError(
                "This version of PowerPoint exposes no animation timeline "
                "(Slide.TimeLine)"
            ) from exc

        animation = sequence.AddEffect(
            Shape=shape,
            effectId=effect_constant,
            Level=level_constant,
            trigger=trigger_constant,
        )

        if exit_effect:
            animation.Exit = MSO_TRUE

        applied: dict[str, Any] = {}
        timing = animation.Timing
        if duration is not None:
            try:
                timing.Duration = float(duration)
                applied["duration"] = float(duration)
            except com_error:
                applied["duration"] = None
        if delay is not None:
            timing.TriggerDelayTime = float(delay)
            applied["delay"] = float(delay)

        self._goto_slide(int(slide_index))
        return {
            "slide_index": int(slide_index),
            "shape_id": int(shape.Id),
            "shape_name": to_python(shape.Name),
            "effect": effect,
            "trigger": trigger,
            "level": level,
            "exit_effect": bool(exit_effect),
            "sequence_index": int(sequence.Count),
            **applied,
        }

    @action("list_animations")
    def list_animations(self, slide_index: int) -> dict[str, Any]:
        """Returns the slide's animation sequence in playback order."""
        slide = self.slide(slide_index)

        try:
            sequence = slide.TimeLine.MainSequence
        except (com_error, AttributeError) as exc:
            raise UnsupportedOperationError(
                "This version of PowerPoint exposes no animation timeline "
                "(Slide.TimeLine)"
            ) from exc

        effects: list[dict[str, Any]] = []
        for index in range(1, sequence.Count + 1):
            animation = sequence(index)
            entry: dict[str, Any] = {
                "index": index,
                "effect": constant_name(animation.EffectType, MSO_ANIM_EFFECTS),
                "effect_id": int(animation.EffectType),
                "exit_effect": bool(animation.Exit == MSO_TRUE),
            }
            try:
                entry["shape_name"] = to_python(animation.Shape.Name)
                entry["shape_id"] = int(animation.Shape.Id)
            except com_error:
                entry["shape_name"] = None
                entry["shape_id"] = None
            try:
                timing = animation.Timing
                entry["trigger"] = constant_name(timing.TriggerType, MSO_ANIM_TRIGGERS)
                entry["duration"] = round(float(timing.Duration), 3)
                entry["delay"] = round(float(timing.TriggerDelayTime), 3)
            except com_error:
                pass
            effects.append(entry)

        transition = slide.SlideShowTransition
        return {
            "slide_index": int(slide_index),
            "count": len(effects),
            "effects": effects,
            "transition": constant_name(transition.EntryEffect, PP_TRANSITIONS),
        }

    @action("set_transition")
    def set_transition(
        self,
        effect: str = "fade",
        slide_index: int | None = None,
        duration: float | None = None,
        advance_on_click: bool = True,
        advance_after: float | None = None,
    ) -> dict[str, Any]:
        """Sets the slide transition; without ``slide_index`` the whole deck.

        ``advance_after`` w sekundach wlacza automatyczne przejscie po czasie -
        niezaleznie od ``advance_on_click``.
        """
        presentation = self.presentation()
        effect_constant = lookup_constant(effect, PP_TRANSITIONS, "effect")

        if slide_index is None:
            indexes = list(range(1, presentation.Slides.Count + 1))
        else:
            indexes = [self.require_index(slide_index, presentation.Slides.Count, "slide_index")]

        duration_applied: float | None = None
        for index in indexes:
            transition = presentation.Slides(index).SlideShowTransition
            transition.EntryEffect = effect_constant
            transition.AdvanceOnClick = MSO_TRUE if advance_on_click else MSO_FALSE

            if duration is not None:
                try:
                    transition.Duration = float(duration)
                    duration_applied = float(duration)
                except (com_error, AttributeError):
                    duration_applied = None

            if advance_after is None:
                transition.AdvanceOnTime = MSO_FALSE
            else:
                transition.AdvanceOnTime = MSO_TRUE
                transition.AdvanceTime = float(advance_after)

        if len(indexes) == 1:
            self._goto_slide(indexes[0])

        return {
            "slides": indexes,
            "effect": effect,
            "duration": duration_applied,
            "advance_on_click": bool(advance_on_click),
            "advance_after": float(advance_after) if advance_after is not None else None,
        }


def _normalize_outline(items: Any) -> list[tuple[str, int]]:
    """Reduces a list of bullet points to ``(text, level)`` pairs.

    Accepts plain text (``"Point"``), dictionaries (``{"text": ..., "level": 2}``)
    and pairs ``("Point", 2)``. The level is clamped to the range 1-5.
    """
    if not isinstance(items, (list, tuple)) or not items:
        raise ValueError("List 'items' cannot be empty")

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
    return entries


def _paragraph_text(text: Any) -> str:
    """Turns newline characters into PowerPoint's paragraph separator.

    COM treats ``\\n`` as a *soft* line break inside one paragraph - the text
    then looks like several lines, but ``Paragraphs().Count`` returns 1 and
    paragraph formatting (spacing, per-paragraph alignment) has nothing to
    latch onto. The real paragraph separator is ``\\r``.
    """
    return str(text).replace("\r\n", "\r").replace("\n", "\r")


def _unit_fraction(value: Any, label: str) -> float:
    """Waliduje ulamek 0.0-1.0; przyjmuje tez procenty podane jako 0-100."""
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a number") from exc

    if 1 < number <= 100:
        number = number / 100
    if not 0.0 <= number <= 1.0:
        raise ValueError(f"{label} must be within 0.0-1.0 (or 0-100%)")
    return number


def _normalize_series(series_data: Any, expected_length: int) -> list[tuple[str, list[Any]]]:
    """Reduces the various series formats to a list of ``(name, values)`` pairs."""
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
            "series_data must be a dictionary, a list of series, or a list of value lists"
        )

    if not series:
        raise InvalidReferenceError("No data for the chart")

    normalized = []
    for name, values in series:
        padded = list(values) + [None] * (expected_length - len(values))
        normalized.append((name, padded[:expected_length]))
    return normalized


__all__ = ["PowerPointController"]
