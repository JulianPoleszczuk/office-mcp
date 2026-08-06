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


def make_paragraph(text="Tekst", outline_level=10, style="Normalny", start=0, end=10):
    paragraph = MagicMock()
    paragraph.Range.Text = text
    paragraph.Range.Start = start
    paragraph.Range.End = end
    paragraph.OutlineLevel = outline_level
    paragraph.Style.NameLocal = style
    return paragraph


def style_property(reject_strings: bool = True):
    """Property mock symulujacy Worda, ktory odrzuca angielskie nazwy stylow."""
    stored: list = []

    def handler(*args):
        if args:
            value = args[0]
            if reject_strings and isinstance(value, str):
                raise make_com_error(-2147352567, "Nie ma takiego stylu")
            stored.append(value)
            return None
        return stored[-1] if stored else None

    return PropertyMock(side_effect=handler), stored


def make_document(paragraphs=None, path=r"C:\dokumenty\raport.docx", name="raport.docx"):
    document = MagicMock()
    paragraph_list = list(paragraphs or [make_paragraph()])
    document.Paragraphs = com_collection(paragraph_list)
    document.Name = name
    document.Path = path.rsplit("\\", 1)[0] if path else ""
    document.FullName = path or ""
    document.Saved = True
    document.Content.Text = "Tekst dokumentu\r"
    document.Sections = com_collection([MagicMock()])
    document.ComputeStatistics.side_effect = lambda statistic: {
        0: 120,
        2: 3,
        3: 800,
    }[statistic]
    return document


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
        target = tmp_path / "notatka.docx"

        controller.dispatch("create_document", {"path": str(target)})

        app.Documents.Add.assert_called_once_with()
        assert document.SaveAs2.call_args[0] == (str(target), 16)

    def test_create_document_with_template(self, word, tmp_path):
        controller, app, _ = word
        template = tmp_path / "firmowy.dotx"
        template.write_bytes(b"x")

        controller.dispatch(
            "create_document",
            {"path": str(tmp_path / "pismo.docx"), "template": str(template)},
        )

        assert app.Documents.Add.call_args.kwargs["Template"] == str(template)

    def test_create_document_missing_template(self, word, tmp_path):
        controller, *_ = word
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "create_document",
                {"path": str(tmp_path / "a.docx"), "template": str(tmp_path / "brak.dotx")},
            )

    def test_open_document_reuses_open_file(self, word, tmp_path):
        controller, app, document = word
        existing = tmp_path / "raport.docx"
        existing.write_bytes(b"x")
        document.FullName = str(existing)

        result = controller.dispatch("open_document", {"path": str(existing)})

        assert result["already_open"] is True
        app.Documents.Open.assert_not_called()

    def test_open_document_opens_new_file(self, word, tmp_path):
        controller, app, _ = word
        other = tmp_path / "umowa.docx"
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
                make_paragraph("Rozdzial 1\r", outline_level=1, style="Naglowek 1"),
                make_paragraph("Tresc\r", outline_level=10),
                make_paragraph("Podrozdzial\r", outline_level=2, style="Naglowek 2"),
                make_paragraph("   \r", outline_level=1),
            ]
        )

        result = controller.dispatch("get_outline", {})

        assert result["count"] == 2
        assert result["headings"][0] == {
            "paragraph_index": 1,
            "level": 1,
            "text": "Rozdzial 1",
            "style": "Naglowek 1",
        }
        assert result["headings"][1]["level"] == 2


class TestContent:
    def test_add_paragraph_reuses_trailing_empty_paragraph(self, word):
        controller, _app, document = word
        empty = make_paragraph("\r")
        document.Paragraphs = com_collection([empty])

        controller.dispatch("add_paragraph", {"text": "Wstep"})

        document.Paragraphs.Add.assert_not_called()
        assert empty.Range.Text == "Wstep"

    def test_add_paragraph_appends_new_one(self, word):
        controller, _app, document = word
        created = make_paragraph()
        document.Paragraphs.Add.return_value = created

        controller.dispatch("add_paragraph", {"text": "Kolejny akapit"})

        document.Paragraphs.Add.assert_called_once()
        assert created.Range.Text == "Kolejny akapit"

    def test_add_paragraph_applies_style(self, word):
        controller, _app, document = word
        created = make_paragraph()
        document.Paragraphs.Add.return_value = created

        result = controller.dispatch(
            "add_paragraph", {"text": "Cytat", "style": "Quote"}
        )

        assert created.Range.Style == "Quote"
        assert result["style"] == "Quote"

    def test_add_heading_uses_heading_style(self, word):
        controller, _app, document = word
        created = make_paragraph()
        document.Paragraphs.Add.return_value = created

        result = controller.dispatch("add_heading", {"text": "Wstep", "level": 2})

        assert created.Range.Style == "Heading 2"
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
        assert document.Content.Find.Replacement.Text == "Zrobione"
        document.Content.Find.Execute.assert_called_once_with(Replace=2)

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
        document.Content.Text = "Nic tu nie ma"
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
        created = [make_paragraph(start=0, end=5), make_paragraph(start=5, end=12)]
        document.Paragraphs.Add.side_effect = created

        result = controller.dispatch(
            "add_bullet_list",
            {"items": ["Pierwszy", {"text": "Zagniezdzony", "level": 2}]},
        )

        document.Range.assert_called_once_with(0, 12)
        document.Range.return_value.ListFormat.ApplyBulletDefault.assert_called_once()
        assert created[1].Range.ListFormat.ListLevelNumber == 2
        assert result["items"] == 2

    def test_add_numbered_list_applies_numbering(self, word):
        controller, _app, document = word
        document.Paragraphs.Add.side_effect = [make_paragraph(), make_paragraph()]

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
            {"rows": 2, "cols": 2, "data": [["Nazwa", "Ilosc"], ["Sruba", 4]]},
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
        document.Paragraphs.Add.side_effect = make_com_error(-2147417848, "Word zamkniety")

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
        ):
            assert name in actions
