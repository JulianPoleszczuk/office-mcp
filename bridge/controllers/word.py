"""Word controller - text, styles, headings and embedded objects over COM.

Paragraphs are indexed from 1, matching the ``Document.Paragraphs`` collection.
Style names can be given in English (``"Heading 1"``, ``"Normal"``) even in a
localised Word - the controller maps them onto built-in constants.
"""

from __future__ import annotations

import os
from typing import Any

from bridge.controllers.base import BaseController, action, is_connection_error
from bridge.utils.com_helpers import (
    WD_ALIGNMENTS,
    WD_BUILTIN_STYLES,
    WD_SAVE_FORMATS,
    WD_SECTION_BREAKS,
    WD_TABLE_STYLES,
    com_error,
    lookup_constant,
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
WD_EXPORT_FORMAT_PDF = 17
WD_STYLE_NORMAL = -1
WD_BULLET_GALLERY = 1
WD_OUTLINE_NUMBER_GALLERY = 3
WD_LIST_NUMBER_ARABIC = 0
WD_TRAILING_TAB = 0
WD_LIST_APPLY_TO_WHOLE = 0
WD_LIST_BEHAVIOR_MULTILEVEL = 2
WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1
POINTS_PER_LINE = 12.0

WD_LINE_SPACING_RULES: dict[float, int] = {
    1.0: 0,   # wdLineSpaceSingle
    1.5: 1,   # wdLineSpace1pt5
    2.0: 2,   # wdLineSpaceDouble
}
WD_LINE_SPACE_MULTIPLE = 5

# English keys only - these are Word's built-in labels. Any other text is
# treated as a custom label and goes into the document verbatim, which is the
# only reliable way to get localised captions regardless of what Word happens
# to call its built-in labels.
WD_CAPTION_LABELS: dict[str, int] = {
    "figure": -1,
    "table": -2,
    "equation": -3,
}

WD_ORIENTATIONS: dict[str, int] = {
    "portrait": WD_ORIENT_PORTRAIT,
        "landscape": WD_ORIENT_LANDSCAPE,
    }


class WordController(BaseController):
    """``doc_*`` actions - operations on a live Word instance."""

    APP_KEY = "word"
    DISPLAY_NAME = "Word"
    ALERTS_OFF = 0

    def document(self) -> Any:
        """The active document, or a clear error when nothing is open."""
        app = self.app
        if app.Documents.Count == 0:
            raise DocumentNotFoundError(
                "No document open - use doc_create_document or doc_open_document"
            )
        try:
            return app.ActiveDocument
        except com_error:
            return app.Documents(app.Documents.Count)

    def paragraph(self, paragraph_index: Any) -> Any:
        """Paragraph at the given 1-based index, with range validation."""
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

    @staticmethod
    def _inside_paragraph_end(paragraph: Any) -> Any:
        """Insertion point just before the end-of-paragraph mark.

        ``Range.Collapse(wdCollapseEnd)`` lands *after* the paragraph mark, i.e.
        already inside the next paragraph - a footnote or hyperlink inserted that
        way shows up at the start of the following paragraph instead of the end
        """
        target = paragraph.Range
        end = int(target.End)
        target.SetRange(max(int(target.Start), end - 1), max(int(target.Start), end - 1))
        return target

    def _insert_point(self, document: Any, position: Any) -> Any:
        """Where to insert a table: ``start``, ``end`` or a paragraph number.

        A thesis needs its table of contents *after* the title page, not at the
        very start of the file - hence the option to point at a paragraph.
        """
        if isinstance(position, int) or str(position).strip().isdigit():
            index = self.require_index(
                position, int(document.Paragraphs.Count), "position"
            )
            return self._inside_paragraph_end(document.Paragraphs(index))

        if str(position).strip().lower() in ("start", "begin"):
            return document.Range(0, 0)
        return self._end_range(document)

    def _end_range(self, document: Any) -> Any:
        """A range collapsed to the very end of the document."""
        target = document.Content
        target.Collapse(WD_COLLAPSE_END)
        return target

    def _append_paragraph(self, document: Any, text: str) -> Any:
        """Appends a paragraph with text at the end of the document and returns it.

        We deliberately avoid ``Paragraphs.Add`` and assigning to ``Range.Text``:
        the first inserts at the selection, the second overwrites the paragraph
        mark and glues neighbouring paragraphs into one. An empty trailing
        paragraph is reused instead of adding another.
        """
        content = document.Content
        content.Collapse(WD_COLLAPSE_END)

        if self._has_text(document.Paragraphs(document.Paragraphs.Count)):
            content.InsertParagraphAfter()

        content.InsertAfter(str(text))
        return document.Paragraphs(document.Paragraphs.Count)

    @staticmethod
    def _has_text(paragraph: Any) -> bool:
        """Whether a paragraph holds anything beyond paragraph and cell marks."""
        raw = str(paragraph.Range.Text)
        return bool(raw.replace("\r", "").replace("\x07", "").strip())

    def _apply_named_style(self, target: Any, style_name: str) -> str:
        """Applies a style by local name or by Word's built-in constant."""
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
                f"Unknown style '{style_name}'. Use a name from Word or one of: "
                f"{', '.join(sorted(WD_BUILTIN_STYLES))}"
            )

        try:
            target.Style = builtin
        except com_error as exc:
            raise InvalidReferenceError(
                f"Could not apply style '{style_name}'"
            ) from exc
        return wanted

    @action("create_document")
    def create_document(self, path: str, template: str | None = None) -> dict[str, Any]:
        """Creates a document (optionally from a .dotx template) and saves it."""
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
        """Opens the file, or activates it if it is already open."""
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
        """Saves the document, or saves it as a new file."""
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
                "The document has no file yet - pass the path parameter"
            )
        else:
            document.Save()

        return self._document_summary(document)

    @action("close")
    def close(self, save: bool = True) -> dict[str, Any]:
        """Closes the document, optionally saving changes."""
        document = self.document()
        name = str(document.Name)

        if save:
            if not document.Path:
                raise InvalidReferenceError(
                    "The document was never saved - run doc_save with a path first"
                )
            document.Save()

        with self.alerts_suppressed():
            document.Close(SaveChanges=bool(save))

        return {"closed": name, "saved": bool(save)}

    @action("get_document_info")
    def get_document_info(self) -> dict[str, Any]:
        """Document metadata: page and word counts, template, path."""
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
        """Returns the whole document text (paragraphs separated by newlines)."""
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
        """Appends a paragraph at the end of the document, optionally styled."""
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
        """Appends a level 1-9 heading."""
        try:
            heading_level = int(level)
        except (TypeError, ValueError) as exc:
            raise InvalidReferenceError("Heading level must be a number 1-9") from exc

        if not 1 <= heading_level <= 9:
            raise InvalidReferenceError("Heading level must be within 1-9")

        result = self.add_paragraph(text, style=f"Heading {heading_level}")
        result["level"] = heading_level
        return result

    @action("insert_page_break")
    def insert_page_break(self) -> dict[str, Any]:
        """Inserts a hard page break at the end of the document."""
        document = self.document()
        self._end_range(document).InsertBreak(WD_PAGE_BREAK)
        return {"paragraph_count": int(document.Paragraphs.Count)}

    @action("find_replace")
    def find_replace(
        self, old_text: str, new_text: str, match_case: bool = False
    ) -> dict[str, Any]:
        """Replaces text across the document and returns the number of hits.

        All search parameters go in a single ``Execute`` call - setting them as
        properties on the ``Find`` object under late-bound COM reports success
        but does not replace anything.

        With ``match_case=False`` Word matches the case of the inserted text to
        the text it found (same as the Find and Replace dialog).
        """
        if not old_text:
            raise InvalidReferenceError("Parameter old_text cannot be empty")

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
        """Adds a bulleted list (supports nesting levels)."""
        return self._add_list(items, numbered=False)

    @action("add_numbered_list")
    def add_numbered_list(self, items: list[Any]) -> dict[str, Any]:
        """Adds a numbered list (supports nesting levels)."""
        return self._add_list(items, numbered=True)

    def _add_list(self, items: list[Any], numbered: bool) -> dict[str, Any]:
        if not isinstance(items, list) or not items:
            raise InvalidReferenceError("List 'items' cannot be empty")

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

        # We keep paragraph numbers, not COM objects. Every further
        # InsertParagraphAfter przestawia wczesniej pobrane obiekty Paragraph,
        # so a range computed from their Range covered only the last entry -
        # one item got numbered instead of the whole list, and the next list
        # was glued onto the previous one.
        first_index = None
        levels: list[tuple[int, int]] = []
        for text, level in entries:
            self._append_paragraph(document, text)
            index = int(document.Paragraphs.Count)
            levels.append((index, level))
            if first_index is None:
                first_index = index

        last_index = int(document.Paragraphs.Count)
        list_range = document.Range(
            document.Paragraphs(first_index).Range.Start,
            document.Paragraphs(last_index).Range.End,
        )

        nested = any(level > 1 for _, level in levels)
        if nested:
            # Default lists (ApplyNumberDefault/ApplyBulletDefault) are
            # single-level - trying to go deeper fails with OLE error
            # 0x800a1200. Levels only work with a gallery template.
            gallery = WD_OUTLINE_NUMBER_GALLERY if numbered else WD_BULLET_GALLERY
            list_range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=self.app.ListGalleries(gallery).ListTemplates(1),
                ContinuePreviousList=False,
                ApplyTo=WD_LIST_APPLY_TO_WHOLE,
                DefaultListBehavior=WD_LIST_BEHAVIOR_MULTILEVEL,
            )
        elif numbered:
            list_range.ListFormat.ApplyNumberDefault()
        else:
            list_range.ListFormat.ApplyBulletDefault()

        for index, level in levels:
            if level > 1:
                try:
                    document.Paragraphs(index).Range.ListFormat.ListLevelNumber = level
                except com_error:
                    pass

        return {
            "items": len(entries),
            "numbered": bool(numbered),
            "first_paragraph_index": first_index,
            "last_paragraph_index": last_index,
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
        """Formats the font of a whole paragraph."""
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
        """Sets paragraph alignment: left / center / right / justify."""
        key = str(alignment).strip().lower()
        if key not in WD_ALIGNMENTS:
            raise InvalidReferenceError(
                f"Unknown alignment '{alignment}'. Available: "
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
        """Sets page margins; ``unit`` accepts cm, mm, inches or points."""
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
        unit: str = "pt",
        own_paragraph: bool = True,
    ) -> dict[str, Any]:
        """Inserts an image inline (``inline``) or as a floating object (``float``).

        ``width`` and ``height`` default to points, like every other COM
        dimension - ``unit=\"cm\"`` lets you give a human-sized value.
        """
        target_path = self.resolve_existing_path(image_path)
        document = self.document()
        mode = str(position).strip().lower()

        if mode in ("inline", "inline_text"):
            if own_paragraph:
                # Without its own paragraph the image is glued onto the last
                # sentence, which justification then stretches across the page.
                self._append_paragraph(document, "")
            shape = document.InlineShapes.AddPicture(
                FileName=target_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=self._end_range(document),
            )
        elif mode in ("float", "floating", "floating_object"):
            shape = document.Shapes.AddPicture(
                FileName=target_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=50,
                Top=50,
            )
        else:
            raise InvalidReferenceError(
                f"Unknown image position '{position}' - use 'inline' or 'float'"
            )

        if width is not None:
            shape.Width = points(width, unit)
        if height is not None:
            shape.Height = float(height)

        return {
            "image_path": target_path,
            "position": mode,
            "width": round(float(shape.Width), 2),
            "height": round(float(shape.Height), 2),
        }

    def _apply_paragraph_format(
        self,
        target: Any,
        line_spacing: float | None,
        space_before: float | None,
        space_after: float | None,
        first_line_indent: float | None,
        left_indent: float | None,
        right_indent: float | None,
        alignment: str | None,
        keep_with_next: bool | None,
        page_break_before: bool | None,
        widow_control: bool | None,
        unit: str,
    ) -> dict[str, Any]:
        """Sets ``ParagraphFormat`` fields on a paragraph or on a style."""
        applied: dict[str, Any] = {}

        if line_spacing is not None:
            value = float(line_spacing)
            rule = WD_LINE_SPACING_RULES.get(value)
            if rule is None:
                # Poza 1.0 / 1.5 / 2.0 Word oczekuje reguly "wielokrotnosc"
                # and line spacing given in points, where one line = 12 pt.
                target.LineSpacingRule = WD_LINE_SPACE_MULTIPLE
                target.LineSpacing = value * POINTS_PER_LINE
            else:
                target.LineSpacingRule = rule
            applied["line_spacing"] = value

        for name, value, attribute in (
            ("space_before", space_before, "SpaceBefore"),
            ("space_after", space_after, "SpaceAfter"),
            ("first_line_indent", first_line_indent, "FirstLineIndent"),
            ("left_indent", left_indent, "LeftIndent"),
            ("right_indent", right_indent, "RightIndent"),
        ):
            if value is not None:
                setattr(target, attribute, points(value, unit))
                applied[name] = points(value, unit)

        if alignment is not None:
            target.Alignment = lookup_constant(alignment, WD_ALIGNMENTS, "alignment")
            applied["alignment"] = alignment

        for name, value, attribute in (
            ("keep_with_next", keep_with_next, "KeepWithNext"),
            ("page_break_before", page_break_before, "PageBreakBefore"),
            ("widow_control", widow_control, "WidowControl"),
        ):
            if value is not None:
                setattr(target, attribute, bool(value))
                applied[name] = bool(value)

        return applied

    @action("set_paragraph_format")
    def set_paragraph_format(
        self,
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

        Scope is chosen one of three ways: ``style`` changes a style definition
        (e.g. all body text at once through ``"Normal"``), ``paragraph_index``
        with ``count`` covers specific paragraphs, and neither means every
        paragraph in the document or, with ``body_text_only=True``, only body
        z pominieciem naglowkow.
        """
        if all(
            value is None
            for value in (
                line_spacing, space_before, space_after, first_line_indent,
                left_indent, right_indent, alignment, keep_with_next,
                page_break_before, widow_control,
            )
        ):
            raise InvalidReferenceError("No field given to change")

        document = self.document()

        if style is not None:
            resolved = self._resolve_style(document, style)
            applied = self._apply_paragraph_format(
                resolved.ParagraphFormat, line_spacing, space_before, space_after,
                first_line_indent, left_indent, right_indent, alignment,
                keep_with_next, page_break_before, widow_control, unit,
            )
            return {
                "scope": "style",
                "style": to_python(resolved.NameLocal),
                "applied": applied,
            }

        total = int(document.Paragraphs.Count)
        if paragraph_index is not None:
            first = self.require_index(paragraph_index, total, "paragraph_index")
            indexes = list(range(first, min(total, first + max(1, int(count)) - 1) + 1))
        else:
            indexes = list(range(1, total + 1))

        touched = 0
        applied: dict[str, Any] = {}
        for index in indexes:
            paragraph = document.Paragraphs(index)
            if body_text_only and int(paragraph.OutlineLevel) != WD_OUTLINE_BODY_TEXT:
                continue
            applied = self._apply_paragraph_format(
                paragraph, line_spacing, space_before, space_after,
                first_line_indent, left_indent, right_indent, alignment,
                keep_with_next, page_break_before, widow_control, unit,
            )
            touched += 1

        return {
            "scope": "paragraphs",
            "paragraphs": touched,
            "body_text_only": bool(body_text_only),
            "applied": applied,
        }

    def _resolve_style(self, document: Any, style: Any) -> Any:
        """Style object by local name, English name or built-in constant."""
        if isinstance(style, int):
            return document.Styles(style)

        wanted = str(style).strip()
        try:
            return document.Styles(wanted)
        except com_error:
            pass

        builtin = WD_BUILTIN_STYLES.get(wanted.lower())
        if builtin is None:
            raise InvalidReferenceError(
                f"Unknown style '{style}'. Use a name from Word or one of: "
                f"{', '.join(sorted(WD_BUILTIN_STYLES))}"
            )
        return document.Styles(builtin)

    @action("set_heading_numbering")
    def set_heading_numbering(
        self, enable: bool = True, levels: int = 3, indent: float = 0.0
    ) -> dict[str, Any]:
        """Wlacza numeracje rozdzialow 1., 1.1, 1.1.1 powiazana ze stylami naglowkow.

        The scheme is built by hand, level by level, instead of taking a gallery
        preset - gallery templates differ between installations and can produce
        dac numeracje prawnicza ("Artykul I.", "Sekcja 2.01").

        Only paragraphs with a heading style get numbered; body text is left
        pozostaje nietkniety.
        """
        document = self.document()
        depth = max(1, min(int(levels), 9))

        if not enable:
            removed = 0
            for index in range(1, int(document.Paragraphs.Count) + 1):
                paragraph = document.Paragraphs(index)
                if 1 <= int(paragraph.OutlineLevel) <= 9:
                    paragraph.Range.ListFormat.RemoveNumbers()
                    removed += 1
            return {"enabled": False, "headings": removed}

        template = self.app.ListGalleries(WD_OUTLINE_NUMBER_GALLERY).ListTemplates(5)
        for level in range(1, depth + 1):
            list_level = template.ListLevels(level)
            list_level.NumberStyle = WD_LIST_NUMBER_ARABIC
            # "%1." for a chapter, "%1.%2" for a section, and so on.
            list_level.NumberFormat = ".".join(
                f"%{position}" for position in range(1, level + 1)
            ) + ("." if level == 1 else "")
            list_level.TrailingCharacter = WD_TRAILING_TAB
            list_level.StartAt = 1
            list_level.NumberPosition = points(indent, "pt")
            list_level.TextPosition = points(indent, "pt") + points(level * 10, "pt")
            list_level.TabPosition = points(indent, "pt") + points(level * 10, "pt")
            list_level.LinkedStyle = to_python(
                document.Styles(WD_BUILTIN_STYLES[f"heading {level}"]).NameLocal
            )

        numbered = 0
        for index in range(1, int(document.Paragraphs.Count) + 1):
            paragraph = document.Paragraphs(index)
            level = int(paragraph.OutlineLevel)
            if 1 <= level <= depth:
                paragraph.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate=template,
                    ContinuePreviousList=True,
                    ApplyTo=WD_LIST_APPLY_TO_WHOLE,
                    DefaultListBehavior=WD_LIST_BEHAVIOR_MULTILEVEL,
                    ApplyLevel=level,
                )
                numbered += 1

        return {"enabled": True, "levels": depth, "headings": numbered}

    @action("add_caption")
    def add_caption(
        self,
        paragraph_index: int,
        text: str,
        label: str = "figure",
        above: bool = False,
    ) -> dict[str, Any]:
        """Adds a numbered caption next to the given paragraph.

        ``label`` przyjmuje etykiete wbudowana (``figure``, ``table``,
        ``equation``) or any custom text, e.g. ``"Figure"`` - a custom label is
        added to Word's label list when needed.

        Numbering is a Word field, so inserting further captions renumbers the
        earlier ones - after changes it is worth calling
        ``doc_update_fields``.
        """
        document = self.document()
        total = int(document.Paragraphs.Count)
        index = self.require_index(paragraph_index, total, "paragraph_index")

        # Etykieta wbudowana ("figure") idzie jako stala - Word sam dobiera
        # its wording, which depends on the document language and can be
        # "Figure" one time and something else another. Any other text is
        # treated as a custom label and added to Word's label list, so
        # czemu praca po polsku dostaje "Rysunek" niezaleznie od ustawien.
        key = str(label).strip().lower()
        if key in WD_CAPTION_LABELS:
            label_name: Any = WD_CAPTION_LABELS[key]
        else:
            label_name = str(label).strip()
            existing = {
                str(self.app.CaptionLabels(position).Name)
                for position in range(1, int(self.app.CaptionLabels.Count) + 1)
            }
            if label_name not in existing:
                self.app.CaptionLabels.Add(label_name)
        title = str(text)
        if title and not title.startswith((":", ".", " ")):
            title = f": {title}"

        document.Paragraphs(index).Range.InsertCaption(
            Label=label_name,
            Title=title,
            Position=0 if above else 1,
            ExcludeLabel=0,
        )

        return {
            "paragraph_index": index,
            "label": label,
            "label_name": to_python(label_name),
            "text": str(text),
            "above": bool(above),
            "paragraph_count": int(document.Paragraphs.Count),
        }

    @action("insert_table_of_figures")
    def insert_table_of_figures(
        self, label: str = "figure", position: Any = "end"
    ) -> dict[str, Any]:
        """Inserts a table of figures or tables built from captions.

        ``position`` as in the table of contents: ``start``, ``end`` or a number.
        """
        document = self.document()
        # The table is built by label name, so for a built-in one we first have
        # najpierw odczytac z Worda, a wlasna ("Rysunek") bierzemy doslownie -
        # exactly as in add_caption, so both tools see the same list.
        key = str(label).strip().lower()
        if key in WD_CAPTION_LABELS:
            caption_name = to_python(self.app.CaptionLabels(WD_CAPTION_LABELS[key]).Name)
        else:
            caption_name = str(label).strip()
        target = self._insert_point(document, position)

        table = document.TablesOfFigures.Add(Range=target, Caption=caption_name)

        return {
            "label": label,
            "caption_label": caption_name,
            "position": position,
            "entries": len(str(table.Range.Text).strip().splitlines()),
        }

    @action("update_fields")
    def update_fields(self) -> dict[str, Any]:
        """Odswieza pola: spis tresci, spisy rysunkow, numeracje podpisow.

        A table of contents inserted before the chapters are written stays empty
        until refreshed - without this step the document looks broken.
        """
        document = self.document()
        document.Fields.Update()

        for collection in (document.TablesOfContents, document.TablesOfFigures):
            for index in range(1, int(collection.Count) + 1):
                try:
                    collection(index).Update()
                except com_error:
                    pass

        return {
            "fields": int(document.Fields.Count),
            "tables_of_contents": int(document.TablesOfContents.Count),
            "tables_of_figures": int(document.TablesOfFigures.Count),
        }

    @action("set_page_setup")
    def set_page_setup(
        self,
        orientation: str | None = None,
        gutter: float | None = None,
        mirror_margins: bool | None = None,
        different_first_page: bool | None = None,
        section: int | None = None,
        unit: str = "cm",
    ) -> dict[str, Any]:
        """Orientacja, margines na oprawe i marginesy lustrzane - druk dwustronny.

        ``gutter`` is the extra room at the binding edge, and
        ``mirror_margins`` przenosi go na przemian raz w lewo, raz w prawo,
        as in a bound thesis.
        """
        if all(
            value is None
            for value in (orientation, gutter, mirror_margins, different_first_page)
        ):
            raise InvalidReferenceError("No field given to change")

        document = self.document()
        if section is None:
            setups = [document.PageSetup]
            scope = "document"
        else:
            index = self.require_index(section, int(document.Sections.Count), "section")
            setups = [document.Sections(index).PageSetup]
            scope = f"section {index}"

        applied: dict[str, Any] = {}
        for setup in setups:
            if orientation is not None:
                setup.Orientation = lookup_constant(
                    orientation, WD_ORIENTATIONS, "orientation"
                )
                applied["orientation"] = orientation
            if gutter is not None:
                setup.Gutter = points(gutter, unit)
                applied["gutter"] = round(points(gutter, unit), 2)
            if mirror_margins is not None:
                setup.MirrorMargins = bool(mirror_margins)
                applied["mirror_margins"] = bool(mirror_margins)
            if different_first_page is not None:
                setup.DifferentFirstPageHeaderFooter = bool(different_first_page)
                applied["different_first_page"] = bool(different_first_page)

        return {"scope": scope, "applied": applied}

    @action("export_pdf")
    def export_pdf(self, path: str, open_after: bool = False) -> dict[str, Any]:
        """Exports the document to PDF without changing the current file.

        Word - w odroznieniu od PowerPointa - wystawia ``ExportAsFixedFormat``
        in a form pywin32 can call, so no workaround is needed.
        """
        document = self.document()
        target = self.resolve_target_path(path)

        with self.alerts_suppressed():
            document.ExportAsFixedFormat(target, WD_EXPORT_FORMAT_PDF, bool(open_after))

        return {
            "path": target,
            "pages": int(document.ComputeStatistics(WD_STATISTIC_PAGES)),
            "size_bytes": os.path.getsize(target) if os.path.isfile(target) else None,
        }

    @action("get_paragraph")
    def get_paragraph(self, paragraph_index: int, count: int = 1) -> dict[str, Any]:
        """Reads paragraphs with style and alignment - no guessing from text."""
        document = self.document()
        total = int(document.Paragraphs.Count)
        first = self.require_index(paragraph_index, total, "paragraph_index")
        last = min(total, first + max(1, int(count)) - 1)

        paragraphs: list[dict[str, Any]] = []
        for index in range(first, last + 1):
            paragraph = document.Paragraphs(index)
            entry: dict[str, Any] = {
                "index": index,
                "text": to_python(paragraph.Range.Text).rstrip("\r\x07"),
            }
            try:
                entry["style"] = to_python(paragraph.Style.NameLocal)
            except com_error:
                entry["style"] = None
            try:
                entry["outline_level"] = int(paragraph.OutlineLevel)
                entry["alignment"] = int(paragraph.Alignment)
            except com_error:
                pass
            paragraphs.append(entry)

        return {
            "paragraph_count": total,
            "returned": len(paragraphs),
            "paragraphs": paragraphs,
        }

    @action("delete_paragraph")
    def delete_paragraph(self, paragraph_index: int, count: int = 1) -> dict[str, Any]:
        """Deletes a paragraph (or several) - the document is no longer append-only."""
        document = self.document()
        total = int(document.Paragraphs.Count)
        first = self.require_index(paragraph_index, total, "paragraph_index")
        amount = max(1, int(count))

        removed: list[str] = []
        for _ in range(amount):
            if int(document.Paragraphs.Count) < first:
                break
            paragraph = document.Paragraphs(first)
            removed.append(to_python(paragraph.Range.Text).rstrip("\r\x07"))
            paragraph.Range.Delete()

        return {
            "deleted": len(removed),
            "texts": removed,
            "paragraph_count": int(document.Paragraphs.Count),
        }

    @action("insert_paragraph")
    def insert_paragraph(
        self,
        text: str,
        paragraph_index: int | None = None,
        after: bool = False,
        style: str | None = None,
    ) -> dict[str, Any]:
        """Inserts a paragraph at a specific place, not just at the end.

        Without ``paragraph_index`` it behaves like ``add_paragraph``. With an
        index it inserts before the given paragraph, or after it with ``after=True``.
        """
        document = self.document()

        if paragraph_index is None:
            paragraph = self._append_paragraph(document, text)
            position = int(document.Paragraphs.Count)
        else:
            total = int(document.Paragraphs.Count)
            index = self.require_index(paragraph_index, total, "paragraph_index")
            anchor = document.Paragraphs(index).Range

            if after:
                anchor.InsertParagraphAfter()
                position = index + 1
            else:
                anchor.InsertParagraphBefore()
                position = index

            document.Paragraphs(position).Range.InsertAfter(str(text))

        applied_style = None
        if style:
            applied_style = self._apply_named_style(
                document.Paragraphs(position).Range, style
            )

        return {
            "paragraph_index": position,
            "text": str(text),
            "style": applied_style,
            "paragraph_count": int(document.Paragraphs.Count),
        }

    @action("add_hyperlink")
    def add_hyperlink(
        self,
        url: str,
        text: str | None = None,
        paragraph_index: int | None = None,
        tooltip: str | None = None,
    ) -> dict[str, Any]:
        """Inserts a hyperlink; without ``paragraph_index`` it appends at the end."""
        if not url:
            raise InvalidReferenceError("'url' cannot be empty")

        document = self.document()
        if paragraph_index is None:
            anchor = self._end_range(document)
        else:
            total = int(document.Paragraphs.Count)
            index = self.require_index(paragraph_index, total, "paragraph_index")
            anchor = self._inside_paragraph_end(document.Paragraphs(index))

        link = document.Hyperlinks.Add(
            Anchor=anchor,
            Address=str(url),
            ScreenTip=str(tooltip) if tooltip else "",
            TextToDisplay=str(text) if text else str(url),
        )

        return {
            "url": to_python(link.Address),
            "text": str(text) if text else str(url),
            "tooltip": tooltip,
            "paragraph_index": paragraph_index or int(document.Paragraphs.Count),
        }

    @action("add_footnote")
    def add_footnote(self, paragraph_index: int, text: str) -> dict[str, Any]:
        """Adds a footnote at the end of the given paragraph."""
        document = self.document()
        total = int(document.Paragraphs.Count)
        index = self.require_index(paragraph_index, total, "paragraph_index")

        anchor = self._inside_paragraph_end(document.Paragraphs(index))
        footnote = document.Footnotes.Add(Range=anchor, Text=str(text))

        return {
            "paragraph_index": index,
            "footnote_index": int(footnote.Index),
            "text": str(text),
            "footnote_count": int(document.Footnotes.Count),
        }

    @action("insert_section_break")
    def insert_section_break(
        self, break_type: str = "next_page", paragraph_index: int | None = None
    ) -> dict[str, Any]:
        """Podzial sekcji: ``next_page``, ``continuous``, ``even_page``, ``odd_page``."""
        document = self.document()
        constant = lookup_constant(break_type, WD_SECTION_BREAKS, "break_type")

        if paragraph_index is None:
            target = self._end_range(document)
        else:
            total = int(document.Paragraphs.Count)
            index = self.require_index(paragraph_index, total, "paragraph_index")
            target = document.Paragraphs(index).Range
            target.Collapse(WD_COLLAPSE_END)

        target.InsertBreak(constant)

        return {
            "break_type": break_type,
            "section_count": int(document.Sections.Count),
            "paragraph_count": int(document.Paragraphs.Count),
        }

    @action("set_columns")
    def set_columns(
        self, count: int = 1, section: int = 1, spacing: float | None = None
    ) -> dict[str, Any]:
        """Sets the number of text columns in a section (newspaper layout)."""
        document = self.document()
        index = self.require_index(section, int(document.Sections.Count), "section")
        columns = document.Sections(index).PageSetup.TextColumns

        columns.SetCount(max(1, int(count)))
        if spacing is not None:
            columns.Spacing = points(spacing, "pt")

        return {
            "section": index,
            "columns": int(columns.Count),
            "spacing": round(float(columns.Spacing), 2),
        }

    @action("set_default_font")
    def set_default_font(
        self, name: str | None = None, size: float | None = None
    ) -> dict[str, Any]:
        """Changes the Normal style font - the basis of the whole document."""
        if not name and size is None:
            raise InvalidReferenceError("Podaj 'name', 'size' albo oba")

        document = self.document()
        font = document.Styles(WD_STYLE_NORMAL).Font

        if name:
            font.Name = str(name)
        if size is not None:
            font.Size = float(size)

        return {
            "name": to_python(font.Name),
            "size": round(float(font.Size), 1),
        }

    @action("format_table")
    def format_table(
        self,
        table_index: int = 1,
        style: str | None = None,
        borders: bool | None = None,
        header_bold: bool | None = None,
        header_fill: Any = None,
        column_widths: list[float] | None = None,
        autofit: bool | None = None,
    ) -> dict[str, Any]:
        """Formats an inserted table - style, borders, header, widths.

        ``style`` takes language-independent names (``light_grid``,
        ``medium_shading1``, ``colorful_list``...), bo wbudowane style tabel
        Word translates built-in table style names, so assigning the English
        name fails with "the given name does not exist".
        """
        document = self.document()
        total = int(document.Tables.Count)
        if not total:
            raise InvalidReferenceError("The document contains no tables")

        index = self.require_index(table_index, total, "table_index")
        table = document.Tables(index)
        applied: dict[str, Any] = {}

        if style is not None:
            table.Style = lookup_constant(style, WD_TABLE_STYLES, "style")
            applied["style"] = style
        if borders is not None:
            line_style = WD_LINE_STYLE_SINGLE if borders else 0
            table.Borders.InsideLineStyle = line_style
            table.Borders.OutsideLineStyle = line_style
            applied["borders"] = bool(borders)
        if header_bold is not None:
            table.Rows(1).Range.Font.Bold = bool(header_bold)
            applied["header_bold"] = bool(header_bold)
        if header_fill is not None:
            table.Rows(1).Shading.BackgroundPatternColor = parse_color(header_fill)
            applied["header_fill"] = str(header_fill)
        if column_widths:
            for position, width in enumerate(column_widths, start=1):
                if position > int(table.Columns.Count):
                    break
                table.Columns(position).Width = points(width, "pt")
            applied["column_widths"] = len(column_widths)
        if autofit:
            table.AutoFitBehavior(2)  # wdAutoFitWindow
            applied["autofit"] = True

        return {
            "table_index": index,
            "rows": int(table.Rows.Count),
            "columns": int(table.Columns.Count),
            "applied": applied,
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
        """Inserts a table at the end of the document or after a given paragraph."""
        row_count, column_count = int(rows), int(cols)
        if row_count < 1 or column_count < 1:
            raise InvalidReferenceError("A table needs at least 1 row and 1 column")

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
        """Sets the page header text in the chosen section."""
        document = self.document()
        index = self.require_index(section, document.Sections.Count, "section")
        header = document.Sections(index).Headers(WD_HEADER_FOOTER_PRIMARY)
        header.Range.Text = str(text)
        return {"section": index, "header": str(text)}

    @action("insert_footer")
    def insert_footer(self, text: str, section: int = 1) -> dict[str, Any]:
        """Sets the footer text in the chosen section."""
        document = self.document()
        index = self.require_index(section, document.Sections.Count, "section")
        footer = document.Sections(index).Footers(WD_HEADER_FOOTER_PRIMARY)
        footer.Range.Text = str(text)
        return {"section": index, "footer": str(text)}

    @action("add_page_numbers")
    def add_page_numbers(
        self, alignment: str = "center", first_page: bool = True, section: int = 1
    ) -> dict[str, Any]:
        """Inserts page numbers into the footer."""
        key = str(alignment).strip().lower()
        if key not in WD_ALIGNMENTS:
            raise InvalidReferenceError(
                f"Unknown alignment '{alignment}'. Available: "
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
        self, levels: int = 3, position: Any = "start"
    ) -> dict[str, Any]:
        """Inserts a table of contents built from heading styles.

        ``position`` accepts ``start``, ``end`` or a paragraph number - the last
        ostatni pozwala umiescic spis za strona tytulowa.
        """
        depth = max(1, min(int(levels), 9))
        document = self.document()

        target = self._insert_point(document, position)

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
                "Could not insert the table of contents - the document must "
                "contain paragraphs with heading styles"
            ) from exc

        return {"levels": depth, "position": str(position).lower()}


__all__ = ["WordController"]
