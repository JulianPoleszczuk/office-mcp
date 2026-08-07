from unittest.mock import MagicMock, PropertyMock

import pytest

from bridge.controllers.word import WordController
from bridge.utils.errors import (
    ComConnectionError,
    DocumentNotFoundError,
    InvalidReferenceError,
    UnsupportedOperationError,
)
from tests.conftest import FakeConnection, com_collection, make_com_error


def make_paragraph(text="Text", outline_level=10, style="Normal", start=0, end=10):
    paragraph = MagicMock()
    paragraph.Range.Text = text
    paragraph.Range.Start = start
    paragraph.Range.End = end
    paragraph.OutlineLevel = outline_level
    paragraph.Style.NameLocal = style
    return paragraph


def style_property(reject_strings: bool = True):
    """Property mock simulating a Word that rejects English style names."""
    stored: list = []

    def handler(*args):
        if args:
            value = args[0]
            if reject_strings and isinstance(value, str):
                raise make_com_error(-2147352567, "No such style")
            stored.append(value)
            return None
        return stored[-1] if stored else None

    return PropertyMock(side_effect=handler), stored


def make_document(paragraphs=None, path=r"C:\documents\raport.docx", name="report.docx"):
    document = MagicMock()
    paragraph_list = list(paragraphs or [make_paragraph()])
    document.Paragraphs = com_collection(paragraph_list)
    document.Name = name
    document.Path = path.rsplit("\\", 1)[0] if path else ""
    document.FullName = path or ""
    document.Saved = True
    document.Content.Text = "Document text\r"
    document.Sections = com_collection([MagicMock()])
    document.ComputeStatistics.side_effect = lambda statistic: {
        0: 120,
        2: 3,
        3: 800,
    }[statistic]
    return document


def growing_paragraphs(document):
    """Mock paragraph collection that grows the way real Word does.

    A static collection would not catch the bug where the list range covered
    only the last entry - with a fixed ``Count`` the first and last index are
    the same paragraph.
    """
    items = [make_paragraph("\r", start=0, end=1)]

    collection = MagicMock()
    collection.side_effect = lambda index, *_: items[int(index) - 1]
    collection.__iter__ = lambda _self: iter(items)
    type(collection).Count = property(lambda _self: len(items))
    document.Paragraphs = collection

    def new_paragraph(*_args, **_kwargs):
        start = 10 * len(items)
        items.append(make_paragraph("\r", start=start, end=start + 1))

    def write_text(text, *_args, **_kwargs):
        start = items[-1].Range.Start
        items[-1] = make_paragraph(f"{text}\r", start=start, end=start + len(str(text)))

    document.Content.InsertParagraphAfter.side_effect = new_paragraph
    document.Content.InsertAfter.side_effect = write_text
    return items


@pytest.fixture
def word():
    document = make_document()

    app = MagicMock()
    app.Documents = com_collection([document])
    app.ActiveDocument = document
    app.Documents.Add.return_value = document
    app.Documents.Open.return_value = document

    controller = WordController(FakeConnection(app=app, key="word"))
    return controller, app, document


class TestFileOperations:
    def test_no_document_open(self, word):
        controller, app, _ = word
        app.Documents = com_collection([])
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch("get_document_info", {})

    def test_create_document_uses_docx_format(self, word, tmp_path):
        controller, app, document = word
        target = tmp_path / "note.docx"

        controller.dispatch("create_document", {"path": str(target)})

        app.Documents.Add.assert_called_once_with()
        assert document.SaveAs2.call_args[0] == (str(target), 16)

    def test_create_document_with_template(self, word, tmp_path):
        controller, app, _ = word
        template = tmp_path / "company.dotx"
        template.write_bytes(b"x")

        controller.dispatch(
            "create_document",
            {"path": str(tmp_path / "letter.docx"), "template": str(template)},
        )

        assert app.Documents.Add.call_args.kwargs["Template"] == str(template)

    def test_create_document_missing_template(self, word, tmp_path):
        controller, *_ = word
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "create_document",
                {"path": str(tmp_path / "a.docx"), "template": str(tmp_path / "missing.dotx")},
            )

    def test_open_document_reuses_open_file(self, word, tmp_path):
        controller, app, document = word
        existing = tmp_path / "report.docx"
        existing.write_bytes(b"x")
        document.FullName = str(existing)

        result = controller.dispatch("open_document", {"path": str(existing)})

        assert result["already_open"] is True
        app.Documents.Open.assert_not_called()

    def test_open_document_opens_new_file(self, word, tmp_path):
        controller, app, _ = word
        other = tmp_path / "contract.docx"
        other.write_bytes(b"x")

        result = controller.dispatch("open_document", {"path": str(other)})

        assert result["already_open"] is False
        app.Documents.Open.assert_called_once_with(str(other))

    def test_save_requires_path_for_new_document(self, word):
        controller, _app, document = word
        document.Path = ""
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("save", {})

    def test_close_without_save(self, word):
        controller, _app, document = word

        result = controller.dispatch("close", {"save": False})

        document.Save.assert_not_called()
        document.Close.assert_called_once_with(SaveChanges=False)
        assert result["saved"] is False


class TestInspection:
    def test_document_info_includes_statistics(self, word):
        controller, *_ = word
        info = controller.dispatch("get_document_info", {})

        assert info["pages"] == 3
        assert info["words"] == 120
        assert info["characters"] == 800

    def test_get_full_text_normalizes_line_breaks(self, word):
        controller, _app, document = word
        document.Content.Text = "Naglowek\rTresc\r\x07"

        result = controller.dispatch("get_full_text", {})

        assert result["text"] == "Naglowek\nTresc\n"

    def test_get_outline_returns_headings_only(self, word):
        controller, _app, document = word
        document.Paragraphs = com_collection(
            [
                make_paragraph("Chapter 1\r", outline_level=1, style="Heading 1"),
                make_paragraph("Body\r", outline_level=10),
                make_paragraph("Subsection\r", outline_level=2, style="Naglowek 2"),
                make_paragraph("   \r", outline_level=1),
            ]
        )

        result = controller.dispatch("get_outline", {})

        assert result["count"] == 2
        assert result["headings"][0] == {
            "paragraph_index": 1,
            "level": 1,
            "text": "Chapter 1",
            "style": "Heading 1",
        }
        assert result["headings"][1]["level"] == 2


class TestContent:
    def test_add_paragraph_reuses_trailing_empty_paragraph(self, word):
        controller, _app, document = word
        document.Paragraphs = com_collection([make_paragraph("\r")])

        controller.dispatch("add_paragraph", {"text": "Intro"})

        document.Content.InsertParagraphAfter.assert_not_called()
        document.Content.InsertAfter.assert_called_once_with("Intro")

    def test_add_paragraph_starts_new_one_after_filled_paragraph(self, word):
        controller, _app, document = word

        controller.dispatch("add_paragraph", {"text": "Another paragraph"})

        document.Content.Collapse.assert_called_once_with(0)
        document.Content.InsertParagraphAfter.assert_called_once()
        document.Content.InsertAfter.assert_called_once_with("Another paragraph")

    def test_add_paragraph_keeps_paragraph_mark(self, word):
        controller, _app, document = word

        controller.dispatch("add_paragraph", {"text": "Tresc"})

        assert document.Paragraphs(1).Range.Text == "Text"

    def test_add_paragraph_applies_style(self, word):
        controller, _app, document = word

        result = controller.dispatch(
            "add_paragraph", {"text": "Quote", "style": "Quote"}
        )

        assert document.Paragraphs(1).Range.Style == "Quote"
        assert result["style"] == "Quote"

    def test_add_heading_uses_heading_style(self, word):
        controller, _app, document = word

        result = controller.dispatch("add_heading", {"text": "Intro", "level": 2})

        assert document.Paragraphs(1).Range.Style == "Heading 2"
        assert result["level"] == 2

    @pytest.mark.parametrize("level", [0, 10, "drugi"])
    def test_add_heading_rejects_bad_level(self, word, level):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_heading", {"text": "X", "level": level})

    def test_insert_page_break(self, word):
        controller, _app, document = word
        controller.dispatch("insert_page_break", {})
        document.Content.InsertBreak.assert_called_once_with(7)

    def test_find_replace_counts_occurrences(self, word):
        controller, _app, document = word
        document.Content.Text = "TODO pierwsze, todo drugie"
        document.Content.Find.Execute.return_value = True

        result = controller.dispatch(
            "find_replace", {"old_text": "TODO", "new_text": "Zrobione"}
        )

        assert result["replacements"] == 2
        kwargs = document.Content.Find.Execute.call_args.kwargs
        assert kwargs["FindText"] == "TODO"
        assert kwargs["ReplaceWith"] == "Zrobione"
        assert kwargs["Replace"] == 2
        assert kwargs["MatchCase"] is False

    def test_find_replace_respects_match_case(self, word):
        controller, _app, document = word
        document.Content.Text = "TODO pierwsze, todo drugie"
        document.Content.Find.Execute.return_value = True

        result = controller.dispatch(
            "find_replace",
            {"old_text": "TODO", "new_text": "Zrobione", "match_case": True},
        )

        assert result["replacements"] == 1

    def test_find_replace_without_hits(self, word):
        controller, _app, document = word
        document.Content.Text = "Nothing here"
        document.Content.Find.Execute.return_value = False

        result = controller.dispatch(
            "find_replace", {"old_text": "TODO", "new_text": "Zrobione"}
        )

        assert result["replacements"] == 0

    def test_find_replace_rejects_empty_needle(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("find_replace", {"old_text": "", "new_text": "x"})

    def test_add_bullet_list_applies_bullets(self, word):
        controller, _app, document = word
        items = growing_paragraphs(document)

        result = controller.dispatch(
            "add_bullet_list",
            {"items": ["First", {"text": "Nested", "level": 2}]},
        )

        assert document.Content.InsertAfter.call_count == 2
        # The range covers both entries - not only the last one.
        document.Range.assert_called_once_with(items[0].Range.Start, items[-1].Range.End)
        assert items[1].Range.ListFormat.ListLevelNumber == 2
        assert result["items"] == 2

    def test_add_numbered_list_applies_numbering(self, word):
        controller, _app, document = word

        controller.dispatch("add_numbered_list", {"items": ["Raz", "Dwa"]})

        document.Range.return_value.ListFormat.ApplyNumberDefault.assert_called_once()

    def test_lists_reject_empty_input(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_bullet_list", {"items": []})


class TestFormatting:
    def test_set_text_style_applies_selected_properties(self, word):
        controller, _app, document = word
        paragraph = document.Paragraphs(1)

        result = controller.dispatch(
            "set_text_style",
            {"paragraph_index": 1, "font_size": 14, "color": "#0000FF", "bold": True},
        )

        assert paragraph.Range.Font.Size == 14.0
        assert paragraph.Range.Font.Color == 0xFF0000
        assert result["applied"]["bold"] is True

    def test_paragraph_index_out_of_range(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_text_style", {"paragraph_index": 9, "bold": True})

    def test_set_paragraph_alignment(self, word):
        controller, _app, document = word

        controller.dispatch(
            "set_paragraph_alignment", {"paragraph_index": 1, "alignment": "center"}
        )

        assert document.Paragraphs(1).Alignment == 1

    def test_set_paragraph_alignment_rejects_unknown(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_paragraph_alignment", {"paragraph_index": 1, "alignment": "ukosnie"}
            )

    def test_apply_style_falls_back_to_builtin_constant(self, word):
        controller, _app, document = word
        paragraph = document.Paragraphs(1)
        prop, stored = style_property()
        type(paragraph.Range).Style = prop

        result = controller.dispatch(
            "apply_style", {"paragraph_index": 1, "style_name": "Heading 1"}
        )

        assert stored == [-2]
        assert result["style"] == "Heading 1"

    def test_apply_style_rejects_unknown_name(self, word):
        controller, _app, document = word
        paragraph = document.Paragraphs(1)
        prop, _ = style_property()
        type(paragraph.Range).Style = prop

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "apply_style", {"paragraph_index": 1, "style_name": "Fikusny"}
            )

    def test_set_page_margins_converts_centimeters(self, word):
        controller, _app, document = word

        result = controller.dispatch(
            "set_page_margins",
            {"top": 2, "bottom": 2, "left": 2.5, "right": 2.5, "unit": "cm"},
        )

        assert round(document.PageSetup.TopMargin, 1) == 56.7
        assert result["margins_points"]["left"] == 70.87

    def test_set_page_margins_rejects_unknown_unit(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_page_margins",
                {"top": 1, "bottom": 1, "left": 1, "right": 1, "unit": "lokcie"},
            )


class TestObjects:
    def test_insert_image_inline(self, word, tmp_path):
        controller, _app, document = word
        image = tmp_path / "logo.png"
        image.write_bytes(b"png")

        result = controller.dispatch("insert_image", {"image_path": str(image)})

        kwargs = document.InlineShapes.AddPicture.call_args.kwargs
        assert kwargs["FileName"] == str(image)
        assert result["position"] == "inline"

    def test_insert_image_float_sets_size(self, word, tmp_path):
        controller, _app, document = word
        image = tmp_path / "logo.png"
        image.write_bytes(b"png")

        controller.dispatch(
            "insert_image",
            {"image_path": str(image), "position": "float", "width": 200, "height": 100},
        )

        shape = document.Shapes.AddPicture.return_value
        assert shape.Width == 200.0
        assert shape.Height == 100.0

    def test_insert_image_rejects_unknown_position(self, word, tmp_path):
        controller, *_ = word
        image = tmp_path / "logo.png"
        image.write_bytes(b"png")

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "insert_image", {"image_path": str(image), "position": "obok"}
            )

    def test_insert_image_missing_file(self, word, tmp_path):
        controller, *_ = word
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch("insert_image", {"image_path": str(tmp_path / "brak.png")})

    def test_insert_table_fills_cells(self, word):
        controller, _app, document = word

        result = controller.dispatch(
            "insert_table",
            {"rows": 2, "cols": 2, "data": [["Name", "Quantity"], ["Screw", 4]]},
        )

        table = document.Tables.Add.return_value
        assert result["cells_filled"] == 4
        table.Cell.assert_any_call(2, 2)

    def test_insert_table_after_paragraph(self, word):
        controller, _app, document = word

        controller.dispatch("insert_table", {"rows": 1, "cols": 1, "position": 1})

        target = document.Paragraphs(1).Range
        target.Collapse.assert_called_once_with(0)
        document.Tables.Add.assert_called_once_with(target, 1, 1)

    def test_insert_table_rejects_zero_size(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("insert_table", {"rows": 0, "cols": 3})

    def test_insert_header_and_footer(self, word):
        controller, _app, document = word

        controller.dispatch("insert_header", {"text": "Firma sp. z o.o."})
        controller.dispatch("insert_footer", {"text": "Poufne"})

        section = document.Sections(1)
        assert section.Headers.return_value.Range.Text == "Firma sp. z o.o."
        assert section.Footers.return_value.Range.Text == "Poufne"

    def test_add_page_numbers(self, word):
        controller, _app, document = word

        controller.dispatch("add_page_numbers", {"alignment": "right"})

        footer = document.Sections(1).Footers.return_value
        footer.PageNumbers.Add.assert_called_once_with(
            PageNumberAlignment=2, FirstPage=True
        )

    def test_add_page_numbers_rejects_unknown_alignment(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_page_numbers", {"alignment": "gdzies"})

    def test_insert_table_of_contents(self, word):
        controller, _app, document = word

        result = controller.dispatch("insert_table_of_contents", {"levels": 2})

        kwargs = document.TablesOfContents.Add.call_args.kwargs
        assert kwargs["LowerHeadingLevel"] == 2
        assert kwargs["UseHeadingStyles"] is True
        document.TablesOfContents.Add.return_value.Update.assert_called_once()
        assert result["levels"] == 2

    def test_insert_table_of_contents_without_headings(self, word):
        controller, _app, document = word
        document.TablesOfContents.Add.side_effect = make_com_error(
            -2147352567, "brak naglowkow"
        )

        with pytest.raises(UnsupportedOperationError):
            controller.dispatch("insert_table_of_contents", {})


class TestErrorMapping:
    def test_disconnected_word(self, word):
        controller, _app, document = word
        document.Content.InsertAfter.side_effect = make_com_error(
            -2147417848, "Word zamkniety"
        )

        with pytest.raises(ComConnectionError):
            controller.dispatch("add_paragraph", {"text": "x"})

    def test_member_not_found_is_unsupported(self, word):
        controller, _app, document = word
        document.TablesOfContents.Add.side_effect = make_com_error(
            -2147352573, "brak metody"
        )

        with pytest.raises(UnsupportedOperationError):
            controller.dispatch("insert_table_of_contents", {})

    def test_actions_cover_public_api(self):
        actions = WordController.actions()
        for name in (
            "create_document",
            "open_document",
            "save",
            "close",
            "get_document_info",
            "get_full_text",
            "get_outline",
            "add_paragraph",
            "add_heading",
            "insert_page_break",
            "find_replace",
            "add_bullet_list",
            "add_numbered_list",
            "set_text_style",
            "set_paragraph_alignment",
            "apply_style",
            "set_page_margins",
            "insert_image",
            "insert_table",
            "insert_header",
            "insert_footer",
            "add_page_numbers",
            "insert_table_of_contents",
            "export_pdf",
            "get_paragraph",
            "delete_paragraph",
            "insert_paragraph",
            "add_hyperlink",
            "add_footnote",
            "insert_section_break",
            "set_columns",
            "set_default_font",
            "format_table",
        ):
            assert name in actions


class TestImageAndCaption:
    def test_image_width_respects_unit(self, word):
        controller, _app, document = word
        shape = MagicMock()
        document.InlineShapes.AddPicture.return_value = shape

        controller.dispatch(
            "insert_image", {"image_path": __file__, "width": 12, "unit": "cm"}
        )

        assert shape.Width == pytest.approx(12 * 28.3464567)

    def test_image_width_defaults_to_points(self, word):
        controller, _app, document = word
        shape = MagicMock()
        document.InlineShapes.AddPicture.return_value = shape

        controller.dispatch("insert_image", {"image_path": __file__, "width": 200})

        assert shape.Width == pytest.approx(200)

    def test_builtin_caption_label_uses_constant(self, word):
        controller, app, document = word
        document.Paragraphs = com_collection([make_paragraph("Image\r")])

        result = controller.dispatch(
            "add_caption", {"paragraph_index": 1, "text": "Chart", "label": "figure"}
        )

        kwargs = document.Paragraphs(1).Range.InsertCaption.call_args.kwargs
        assert kwargs["Label"] == -1  # wdCaptionFigure
        assert kwargs["Title"] == ": Chart"
        assert result["label_name"] == -1
        app.CaptionLabels.Add.assert_not_called()

    def test_custom_caption_label_is_registered_once(self, word):
        controller, app, document = word
        builtin = MagicMock()
        builtin.Name = "Figure"
        app.CaptionLabels = com_collection([builtin])
        document.Paragraphs = com_collection([make_paragraph("Image\r")])

        controller.dispatch(
            "add_caption", {"paragraph_index": 1, "text": "Chart", "label": "Diagram"}
        )

        # A custom label must land in Word's label list, otherwise InsertCaption
        # would not recognise it.
        app.CaptionLabels.Add.assert_called_once_with("Diagram")
        assert document.Paragraphs(1).Range.InsertCaption.call_args.kwargs["Label"] == "Diagram"

    def test_known_custom_label_is_not_added_twice(self, word):
        controller, app, document = word
        existing = MagicMock()
        existing.Name = "Diagram"
        app.CaptionLabels = com_collection([existing])
        document.Paragraphs = com_collection([make_paragraph("Image\r")])

        controller.dispatch(
            "add_caption", {"paragraph_index": 1, "text": "Chart", "label": "Diagram"}
        )

        app.CaptionLabels.Add.assert_not_called()

    def test_image_gets_its_own_paragraph(self, word):
        controller, _app, document = word
        growing_paragraphs(document)

        controller.dispatch("insert_image", {"image_path": __file__, "width": 200})

        # Without its own paragraph the image would stick to the last sentence.
        document.Content.InsertAfter.assert_called_once_with("")

    def test_caption_above_uses_position_zero(self, word):
        controller, app, document = word
        app.CaptionLabels.return_value.Name = "Table"
        document.Paragraphs = com_collection([make_paragraph("Table\r")])

        controller.dispatch(
            "add_caption",
            {"paragraph_index": 1, "text": "Wyniki", "label": "table", "above": True},
        )

        assert document.Paragraphs(1).Range.InsertCaption.call_args.kwargs["Position"] == 0


class TestListRange:
    def test_list_range_spans_every_item(self, word):
        controller, _app, document = word
        items = growing_paragraphs(document)

        controller.dispatch("add_numbered_list", {"items": ["A", "B", "C"]})

        # The range must cover the first and last entry, not only the last.
        first, last = items[0], items[-1]
        assert document.Range.call_args[0] == (first.Range.Start, last.Range.End)
        assert len(items) == 3
        document.Range.return_value.ListFormat.ApplyNumberDefault.assert_called_once()

    def test_bullet_list_uses_bullet_default(self, word):
        controller, _app, document = word
        growing_paragraphs(document)

        controller.dispatch("add_bullet_list", {"items": ["A", "B"]})

        document.Range.return_value.ListFormat.ApplyBulletDefault.assert_called_once()

    def test_nested_list_uses_gallery_template(self, word):
        controller, app, document = word
        items = growing_paragraphs(document)

        controller.dispatch(
            "add_numbered_list", {"items": ["A", {"text": "B", "level": 2}]}
        )

        # Default lists are single-level - going deeper needs a template.
        document.Range.return_value.ListFormat.ApplyNumberDefault.assert_not_called()
        assert app.ListGalleries.call_args[0][0] == 3  # wdOutlineNumberGallery
        assert items[1].Range.ListFormat.ListLevelNumber == 2

    def test_nested_bullets_use_bullet_gallery(self, word):
        controller, app, document = word
        growing_paragraphs(document)

        controller.dispatch(
            "add_bullet_list", {"items": ["A", {"text": "B", "level": 2}]}
        )

        assert app.ListGalleries.call_args[0][0] == 1  # wdBulletGallery

    def test_flat_list_keeps_default_numbering(self, word):
        controller, app, document = word
        growing_paragraphs(document)

        controller.dispatch("add_numbered_list", {"items": ["A", "B"]})

        document.Range.return_value.ListFormat.ApplyNumberDefault.assert_called_once()
        app.ListGalleries.assert_not_called()


class TestParagraphEditing:
    def test_delete_paragraph_returns_removed_text(self, word):
        controller, _app, document = word
        document.Paragraphs = com_collection(
            [make_paragraph("First\r"), make_paragraph("Second\r")]
        )

        result = controller.dispatch("delete_paragraph", {"paragraph_index": 1})

        assert result["deleted"] == 1
        assert result["texts"] == ["First"]

    def test_delete_paragraph_out_of_range(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("delete_paragraph", {"paragraph_index": 99})

    def test_insert_paragraph_before_given_index(self, word):
        controller, _app, document = word
        first = make_paragraph("First\r")
        document.Paragraphs = com_collection([first, make_paragraph("Second\r")])

        result = controller.dispatch(
            "insert_paragraph", {"text": "New", "paragraph_index": 1}
        )

        first.Range.InsertParagraphBefore.assert_called_once()
        assert result["paragraph_index"] == 1

    def test_insert_paragraph_after_given_index(self, word):
        controller, _app, document = word
        first = make_paragraph("First\r")
        document.Paragraphs = com_collection([first, make_paragraph("Second\r")])

        result = controller.dispatch(
            "insert_paragraph",
            {"text": "New", "paragraph_index": 1, "after": True},
        )

        first.Range.InsertParagraphAfter.assert_called_once()
        assert result["paragraph_index"] == 2

    def test_get_paragraph_reads_style_and_text(self, word):
        controller, _app, document = word
        document.Paragraphs = com_collection(
            [make_paragraph("Title\r\x07", style="Heading 1", outline_level=1)]
        )

        result = controller.dispatch("get_paragraph", {"paragraph_index": 1})

        assert result["paragraphs"][0]["text"] == "Title"
        assert result["paragraphs"][0]["style"] == "Heading 1"
        assert result["paragraphs"][0]["outline_level"] == 1

    def test_get_paragraph_clamps_count_to_document(self, word):
        controller, _app, document = word
        document.Paragraphs = com_collection(
            [make_paragraph("A\r"), make_paragraph("B\r")]
        )

        result = controller.dispatch(
            "get_paragraph", {"paragraph_index": 2, "count": 10}
        )

        assert result["returned"] == 1


class TestWordExtras:
    def test_export_pdf_uses_fixed_format(self, word, tmp_path):
        controller, _app, document = word
        target = tmp_path / "report.pdf"

        result = controller.dispatch("export_pdf", {"path": str(target)})

        document.ExportAsFixedFormat.assert_called_once_with(str(target), 17, False)
        document.SaveAs.assert_not_called()
        assert result["pages"] == 3

    def test_hyperlink_defaults_text_to_url(self, word):
        controller, _app, document = word

        controller.dispatch("add_hyperlink", {"url": "https://openai.com"})

        assert (
            document.Hyperlinks.Add.call_args.kwargs["TextToDisplay"]
            == "https://openai.com"
        )

    def test_footnote_anchors_before_paragraph_mark(self, word):
        controller, _app, document = word
        paragraph = make_paragraph("Paragraph text\r", start=100, end=115)
        document.Paragraphs = com_collection([paragraph])

        controller.dispatch(
            "add_footnote", {"paragraph_index": 1, "text": "Footnote"}
        )

        # Collapse(wdCollapseEnd) would land on 115, i.e. already in the next
        # paragraph - the mark would appear before its first word.
        paragraph.Range.SetRange.assert_called_once_with(114, 114)
        assert document.Footnotes.Add.call_args.kwargs["Text"] == "Footnote"

    def test_hyperlink_in_paragraph_anchors_before_mark(self, word):
        controller, _app, document = word
        paragraph = make_paragraph("Text\r", start=10, end=17)
        document.Paragraphs = com_collection([paragraph])

        controller.dispatch(
            "add_hyperlink", {"url": "https://openai.com", "paragraph_index": 1}
        )

        paragraph.Range.SetRange.assert_called_once_with(16, 16)

    def test_hyperlink_requires_url(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_hyperlink", {"url": ""})

    def test_section_break_maps_name(self, word):
        controller, _app, document = word

        controller.dispatch("insert_section_break", {"break_type": "continuous"})

        assert document.Content.InsertBreak.call_args[0][0] == 3

    def test_unknown_section_break_is_rejected(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("insert_section_break", {"break_type": "new_sheet"})

    def test_set_default_font_writes_normal_style(self, word):
        controller, _app, document = word

        controller.dispatch("set_default_font", {"name": "Segoe UI", "size": 11})

        assert document.Styles.call_args[0][0] == -1  # wdStyleNormal
        assert document.Styles.return_value.Font.Name == "Segoe UI"

    def test_set_default_font_without_arguments_is_rejected(self, word):
        controller, *_ = word
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_default_font", {})

    def test_table_style_uses_locale_independent_constant(self, word):
        controller, _app, document = word
        table = MagicMock()
        document.Tables = com_collection([table])

        controller.dispatch("format_table", {"table_index": 1, "style": "light_grid"})

        # Word translates built-in table style names - a constant goes in
        assert table.Style == -161

    def test_format_table_without_tables_is_rejected(self, word):
        controller, _app, document = word
        document.Tables = com_collection([])

        with pytest.raises(InvalidReferenceError):
            controller.dispatch("format_table", {"table_index": 1, "borders": True})

    def test_format_table_column_widths_stop_at_column_count(self, word):
        controller, _app, document = word
        table = MagicMock()
        table.Columns.Count = 2
        document.Tables = com_collection([table])

        result = controller.dispatch(
            "format_table", {"table_index": 1, "column_widths": [100, 120, 140]}
        )

        assert result["applied"]["column_widths"] == 3
        assert table.Columns.call_count == 2

    def test_set_columns_updates_section(self, word):
        controller, _app, document = word
        section = MagicMock()
        section.PageSetup.TextColumns.Count = 2
        section.PageSetup.TextColumns.Spacing = 24.0
        document.Sections = com_collection([section])

        result = controller.dispatch("set_columns", {"count": 2, "spacing": 24})

        section.PageSetup.TextColumns.SetCount.assert_called_once_with(2)
        assert result["columns"] == 2
