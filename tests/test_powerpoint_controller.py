from unittest.mock import MagicMock

import pytest

from bridge.controllers.powerpoint import PowerPointController, _normalize_series
from bridge.utils.errors import (
    ComConnectionError,
    DocumentNotFoundError,
    InvalidReferenceError,
    ProtocolError,
)
from tests.conftest import FakeConnection, com_collection, make_com_error, make_shape


def make_slide(shapes=None, has_title=True, title_text="Tytul"):
    slide = MagicMock()
    shape_list = list(shapes or [])
    title_shape = make_shape(shape_id=1, name="Title 1", text=title_text, placeholder_type=1)

    slide.Shapes = com_collection(shape_list)
    slide.Shapes.HasTitle = has_title
    slide.Shapes.Title = title_shape
    slide.Shapes.Placeholders = com_collection(
        [shape for shape in shape_list if shape.Type == 14]
    )
    slide.CustomLayout.Name = "Tytul i zawartosc"
    return slide


def make_presentation(slides=None, path=r"C:\prezentacje\test.pptx", name="test.pptx"):
    presentation = MagicMock()
    presentation.Slides = com_collection(list(slides or []))
    presentation.Name = name
    presentation.Path = path.rsplit("\\", 1)[0] if path else ""
    presentation.FullName = path or ""
    presentation.Saved = True
    presentation.PageSetup.SlideWidth = 960.0
    presentation.PageSetup.SlideHeight = 540.0
    return presentation


@pytest.fixture
def powerpoint():
    slides = [make_slide(shapes=[make_shape(shape_id=5, text="Punkt", placeholder_type=2)])]
    presentation = make_presentation(slides=slides)

    app = MagicMock()
    app.Presentations = com_collection([presentation])
    app.ActivePresentation = presentation
    app.Presentations.Add.return_value = presentation
    app.Presentations.Open.return_value = presentation

    connection = FakeConnection(app=app, key="powerpoint")
    controller = PowerPointController(connection)
    return controller, app, presentation, slides


class TestDispatch:
    def test_unknown_action_is_protocol_error(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(ProtocolError):
            controller.dispatch("nie_ma_takiej_akcji", {})

    def test_missing_required_param_is_protocol_error(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(ProtocolError) as info:
            controller.dispatch("set_title", {"slide_index": 1})
        assert "text" in info.value.message

    def test_unexpected_param_is_protocol_error(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(ProtocolError):
            controller.dispatch("list_slides", {"kolor": "red"})

    def test_registered_actions_cover_public_api(self):
        actions = PowerPointController.actions()
        for name in (
            "create_presentation",
            "open_presentation",
            "save",
            "close",
            "get_presentation_info",
            "get_slide_content",
            "list_slides",
            "add_slide",
            "delete_slide",
            "duplicate_slide",
            "reorder_slide",
            "set_title",
            "add_textbox",
            "add_bullet_list",
            "find_replace_text",
            "set_speaker_notes",
            "set_text_style",
            "apply_theme",
            "set_background",
            "set_slide_layout",
            "add_image",
            "add_chart",
            "add_table",
            "add_shape",
            "add_animation",
            "list_animations",
            "set_transition",
            "delete_shape",
            "set_shape_position",
            "set_shape_order",
            "export_slide",
            "export_pdf",
            "get_theme",
            "set_theme_colors",
            "set_theme_fonts",
            "set_master_background",
            "set_shape_format",
            "set_paragraph_format",
            "format_chart",
            "group_shapes",
            "ungroup_shapes",
            "align_shapes",
            "distribute_shapes",
            "add_hyperlink",
            "set_headers_footers",
            "add_media",
            "add_smartart",
            "list_smartart_layouts",
            "list_sections",
            "add_section",
            "delete_section",
            "slideshow",
            "copy_slide_to",
        ):
            assert name in actions


class TestFileOperations:
    def test_create_presentation_saves_as_pptx(self, powerpoint, tmp_path):
        controller, app, presentation, _ = powerpoint
        target = tmp_path / "raport.pptx"

        result = controller.dispatch("create_presentation", {"path": str(target)})

        app.Presentations.Add.assert_called_once()
        saved_path, saved_format = presentation.SaveAs.call_args[0]
        assert saved_path == str(target)
        assert saved_format == 24
        assert result["slide_count"] == 1

    def test_create_presentation_applies_template(self, powerpoint, tmp_path):
        controller, _app, presentation, _ = powerpoint
        template = tmp_path / "motyw.potx"
        template.write_bytes(b"x")

        controller.dispatch(
            "create_presentation",
            {"path": str(tmp_path / "a.pptx"), "template": str(template)},
        )

        presentation.ApplyTemplate.assert_called_once_with(str(template))

    def test_create_presentation_rejects_missing_directory(self, powerpoint, tmp_path):
        controller, *_ = powerpoint
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "create_presentation", {"path": str(tmp_path / "brak" / "a.pptx")}
            )

    def test_open_presentation_requires_existing_file(self, powerpoint, tmp_path):
        controller, *_ = powerpoint
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch("open_presentation", {"path": str(tmp_path / "brak.pptx")})

    def test_open_presentation_reuses_open_file(self, powerpoint, tmp_path):
        controller, app, presentation, _ = powerpoint
        existing = tmp_path / "otwarta.pptx"
        existing.write_bytes(b"x")
        presentation.FullName = str(existing)

        result = controller.dispatch("open_presentation", {"path": str(existing)})

        assert result["already_open"] is True
        app.Presentations.Open.assert_not_called()

    def test_open_presentation_opens_new_file(self, powerpoint, tmp_path):
        controller, app, _presentation, _ = powerpoint
        other = tmp_path / "inna.pptx"
        other.write_bytes(b"x")

        result = controller.dispatch("open_presentation", {"path": str(other)})

        assert result["already_open"] is False
        app.Presentations.Open.assert_called_once()

    def test_save_without_path_calls_save(self, powerpoint):
        controller, _app, presentation, _ = powerpoint
        controller.dispatch("save", {})
        presentation.Save.assert_called_once()

    def test_save_unsaved_presentation_without_path_fails(self, powerpoint):
        controller, _app, presentation, _ = powerpoint
        presentation.Path = ""

        with pytest.raises(InvalidReferenceError):
            controller.dispatch("save", {})

    def test_close_without_saving_marks_presentation_saved(self, powerpoint):
        controller, _app, presentation, _ = powerpoint

        result = controller.dispatch("close", {"save": False})

        presentation.Save.assert_not_called()
        presentation.Close.assert_called_once()
        assert result["saved"] is False

    def test_no_presentation_open_is_document_not_found(self, powerpoint):
        controller, app, *_ = powerpoint
        app.Presentations = com_collection([])

        with pytest.raises(DocumentNotFoundError):
            controller.dispatch("list_slides", {})


class TestInspection:
    def test_get_presentation_info(self, powerpoint):
        controller, _app, presentation, _ = powerpoint
        presentation.Designs.return_value.Name = "Motyw firmowy"

        info = controller.dispatch("get_presentation_info", {})

        assert info["slide_width"] == 960.0
        assert info["slide_height"] == 540.0
        assert info["theme"] == "Motyw firmowy"

    def test_list_slides_returns_titles_and_layouts(self, powerpoint):
        controller, *_ = powerpoint
        result = controller.dispatch("list_slides", {})

        assert result["slide_count"] == 1
        assert result["slides"][0]["title"] == "Tytul"
        assert result["slides"][0]["layout"] == "Tytul i zawartosc"

    def test_get_slide_content_lists_shapes_and_notes(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].NotesPage.Shapes.Placeholders.return_value.TextFrame.HasText = True
        slides[0].NotesPage.Shapes.Placeholders.return_value.TextFrame.TextRange.Text = (
            "Notatka"
        )

        result = controller.dispatch("get_slide_content", {"slide_index": 1})

        assert result["notes"] == "Notatka"
        assert result["shapes"][0]["text"] == "Punkt"
        assert result["shapes"][0]["shape_id"] == 5

    def test_slide_index_out_of_range(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("get_slide_content", {"slide_index": 7})

    def test_slide_index_must_be_number(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("get_slide_content", {"slide_index": "drugi"})


class TestStructure:
    def test_add_slide_uses_layout_constant(self, powerpoint):
        controller, _app, presentation, _ = powerpoint

        result = controller.dispatch("add_slide", {"layout": "title_content"})

        presentation.Slides.Add.assert_called_once_with(2, 2)
        assert result["slide_index"] == 2

    def test_add_slide_with_explicit_index_and_title(self, powerpoint):
        controller, _app, presentation, _ = powerpoint
        new_slide = make_slide()
        presentation.Slides.Add.return_value = new_slide

        controller.dispatch("add_slide", {"layout": "blank", "index": 1, "title": "Wstep"})

        presentation.Slides.Add.assert_called_once_with(1, 12)
        assert new_slide.Shapes.Title.TextFrame.TextRange.Text == "Wstep"

    def test_add_slide_rejects_unknown_layout(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_slide", {"layout": "kosmiczny"})

    def test_delete_slide(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        controller.dispatch("delete_slide", {"slide_index": 1})
        slides[0].Delete.assert_called_once()

    def test_reorder_slide_calls_move_to(self, powerpoint):
        controller, _app, presentation, slides = powerpoint
        presentation.Slides = com_collection(slides + [make_slide()])

        controller.dispatch("reorder_slide", {"from_index": 2, "to_index": 1})

        presentation.Slides(2).MoveTo.assert_called_once_with(1)


class TestContent:
    def test_set_title_updates_existing_placeholder(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint

        result = controller.dispatch("set_title", {"slide_index": 1, "text": "Nowy tytul"})

        assert slides[0].Shapes.Title.TextFrame.TextRange.Text == "Nowy tytul"
        assert result["created_textbox"] is False

    def test_set_title_creates_textbox_when_layout_has_no_title(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes.HasTitle = False

        result = controller.dispatch("set_title", {"slide_index": 1, "text": "Tytul"})

        slides[0].Shapes.AddTextbox.assert_called_once()
        assert result["created_textbox"] is True

    def test_add_textbox_applies_formatting(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        textbox = make_shape(shape_id=42)
        slides[0].Shapes.AddTextbox.return_value = textbox

        result = controller.dispatch(
            "add_textbox",
            {
                "slide_index": 1,
                "text": "Tresc",
                "left": 50,
                "top": 100,
                "width": 400,
                "height": 80,
                "font_size": 18,
                "bold": True,
                "color": "#FF0000",
            },
        )

        assert textbox.TextFrame.TextRange.Text == "Tresc"
        assert textbox.TextFrame.TextRange.Font.Size == 18.0
        assert textbox.TextFrame.TextRange.Font.Color.RGB == 0x0000FF
        assert result["shape_id"] == 42

    def test_add_bullet_list_sets_indent_levels(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        placeholder = slides[0].Shapes.Placeholders(1)

        result = controller.dispatch(
            "add_bullet_list",
            {
                "slide_index": 1,
                "items": ["Glowny", {"text": "Podpunkt", "level": 2}],
            },
        )

        text_range = placeholder.TextFrame.TextRange
        assert text_range.Text == "Glowny\rPodpunkt"
        assert text_range.Paragraphs(2).IndentLevel == 2
        assert result["items"] == 2

    def test_add_bullet_list_rejects_empty_items(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_bullet_list", {"slide_index": 1, "items": []})

    def test_find_replace_counts_replacements(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        found = MagicMock()
        found.Start = 1
        shape.TextFrame.TextRange.Replace.side_effect = [found, None]

        result = controller.dispatch(
            "find_replace_text", {"old_text": "TODO", "new_text": "Zrobione"}
        )

        assert result["replacements"] == 1
        assert result["slides_scanned"] == 1

    def test_find_replace_rejects_empty_needle(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("find_replace_text", {"old_text": "", "new_text": "x"})

    def test_set_speaker_notes(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint

        controller.dispatch("set_speaker_notes", {"slide_index": 1, "text": "Notatka"})

        notes = slides[0].NotesPage.Shapes.Placeholders.return_value
        assert notes.TextFrame.TextRange.Text == "Notatka"


class TestFormatting:
    def test_set_text_style_applies_only_given_properties(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        result = controller.dispatch(
            "set_text_style",
            {"slide_index": 1, "shape_id": 5, "font_size": 24, "bold": True},
        )

        assert shape.TextFrame.TextRange.Font.Size == 24.0
        assert shape.TextFrame.TextRange.Font.Bold == -1
        assert result["applied"] == {"font_size": 24.0, "bold": True}

    def test_set_text_style_unknown_shape(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_text_style", {"slide_index": 1, "shape_id": 999})

    def test_set_background_solid_color(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        fill = slides[0].Background.Fill
        fill.ForeColor.RGB = 0

        controller.dispatch("set_background", {"slide_index": 1, "color": "blue"})

        fill.Solid.assert_called_once()
        assert fill.ForeColor.RGB == 0xC07000

    def test_set_background_requires_color_or_image(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_background", {"slide_index": 1})

    def test_apply_theme_rejects_unknown_theme(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch("apply_theme", {"theme_name_or_path": "NieMaTakiego"})

    def test_set_slide_layout_matches_custom_layout_by_name(self, powerpoint):
        controller, _app, presentation, slides = powerpoint
        layout = MagicMock()
        layout.Name = "Porownanie"
        presentation.SlideMaster.CustomLayouts = com_collection([layout])

        result = controller.dispatch(
            "set_slide_layout", {"slide_index": 1, "layout_name": "porownanie"}
        )

        assert result["source"] == "custom_layout"
        assert slides[0].CustomLayout == layout

    def test_set_slide_layout_falls_back_to_builtin(self, powerpoint):
        controller, _app, presentation, slides = powerpoint
        presentation.SlideMaster.CustomLayouts = com_collection([])

        result = controller.dispatch(
            "set_slide_layout", {"slide_index": 1, "layout_name": "blank"}
        )

        assert result["source"] == "builtin"
        assert slides[0].Layout == 12


class TestVisuals:
    def test_add_image_keeps_aspect_ratio_without_size(self, powerpoint, tmp_path):
        controller, _app, _presentation, slides = powerpoint
        image = tmp_path / "wykres.png"
        image.write_bytes(b"png")

        controller.dispatch(
            "add_image",
            {"slide_index": 1, "image_path": str(image), "left": 10, "top": 20},
        )

        kwargs = slides[0].Shapes.AddPicture.call_args.kwargs
        assert kwargs["Width"] == -1
        assert kwargs["Height"] == -1

    def test_add_image_requires_existing_file(self, powerpoint, tmp_path):
        controller, *_ = powerpoint
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "add_image",
                {"slide_index": 1, "image_path": str(tmp_path / "brak.png"), "left": 0, "top": 0},
            )

    def test_add_chart_fills_worksheet_and_sets_source(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=11)
        slides[0].Shapes.AddChart2.return_value = shape
        worksheet = shape.Chart.ChartData.Workbook.Worksheets.return_value
        worksheet.Name = "Arkusz1"
        worksheet.ListObjects.Count = 0
        worksheet.Range.return_value.Address.return_value = "$A$1:$B$3"

        result = controller.dispatch(
            "add_chart",
            {
                "slide_index": 1,
                "chart_type": "bar",
                "categories": ["A", "B"],
                "series_data": {"Wyniki": [1, 2]},
                "left": 50,
                "top": 100,
                "width": 400,
                "height": 300,
                "title": "Wyniki",
            },
        )

        slides[0].Shapes.AddChart2.assert_called_once_with(-1, 57, 50.0, 100.0, 400.0, 300.0)
        shape.Chart.SetSourceData.assert_called_once_with("='Arkusz1'!$A$1:$B$3")
        assert result["series"] == ["Wyniki"]
        assert shape.Chart.ChartTitle.Text == "Wyniki"

    def test_add_chart_rejects_empty_categories(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_chart",
                {
                    "slide_index": 1,
                    "chart_type": "bar",
                    "categories": [],
                    "series_data": {"a": [1]},
                    "left": 0,
                    "top": 0,
                    "width": 10,
                    "height": 10,
                },
            )

    def test_add_chart_rejects_unknown_type(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_chart",
                {
                    "slide_index": 1,
                    "chart_type": "spirala",
                    "categories": ["A"],
                    "series_data": {"a": [1]},
                    "left": 0,
                    "top": 0,
                    "width": 10,
                    "height": 10,
                },
            )

    def test_add_table_fills_cells(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=21)
        slides[0].Shapes.AddTable.return_value = shape

        result = controller.dispatch(
            "add_table",
            {
                "slide_index": 1,
                "rows": 2,
                "cols": 2,
                "data": [["Nazwa", "Wynik"], ["Robot", 12]],
                "left": 0,
                "top": 0,
                "width": 400,
                "height": 100,
            },
        )

        assert result["cells_filled"] == 4
        shape.Table.Cell.assert_any_call(2, 2)

    def test_add_table_ignores_extra_data(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes.AddTable.return_value = make_shape(shape_id=22)

        result = controller.dispatch(
            "add_table",
            {
                "slide_index": 1,
                "rows": 1,
                "cols": 1,
                "data": [["A", "B"], ["C"]],
                "left": 0,
                "top": 0,
                "width": 100,
                "height": 100,
            },
        )

        assert result["cells_filled"] == 1

    def test_add_shape_with_fill(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=31)
        slides[0].Shapes.AddShape.return_value = shape

        controller.dispatch(
            "add_shape",
            {
                "slide_index": 1,
                "shape_type": "rounded_rectangle",
                "left": 0,
                "top": 0,
                "width": 100,
                "height": 50,
                "fill_color": (255, 0, 0),
                "text": "Start",
            },
        )

        slides[0].Shapes.AddShape.assert_called_once_with(5, 0.0, 0.0, 100.0, 50.0)
        assert shape.Fill.ForeColor.RGB == 0x0000FF
        assert shape.TextFrame.TextRange.Text == "Start"

    def test_add_shape_can_drop_theme_outline(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=32)
        slides[0].Shapes.AddShape.return_value = shape

        controller.dispatch(
            "add_shape",
            {
                "slide_index": 1,
                "shape_type": "oval",
                "left": 0,
                "top": 0,
                "width": 20,
                "height": 20,
                "fill_color": "#10A37F",
                "line_color": "none",
            },
        )

        assert shape.Line.Visible == 0
        assert shape.Fill.ForeColor.RGB == 0x7FA310

    def test_add_shape_line_color_and_width(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=33)
        slides[0].Shapes.AddShape.return_value = shape

        controller.dispatch(
            "add_shape",
            {
                "slide_index": 1,
                "shape_type": "rectangle",
                "left": 0,
                "top": 0,
                "width": 20,
                "height": 20,
                "fill_color": "none",
                "line_color": "#FFFFFF",
                "line_width": 1.5,
            },
        )

        assert shape.Fill.Visible == 0
        assert shape.Line.Visible == -1
        assert shape.Line.ForeColor.RGB == 0xFFFFFF
        assert shape.Line.Weight == 1.5


class TestErrorMapping:
    def test_com_disconnect_maps_to_connection_error(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes.AddTextbox.side_effect = make_com_error(-2147417848, "brak apki")

        with pytest.raises(ComConnectionError):
            controller.dispatch(
                "add_textbox",
                {
                    "slide_index": 1,
                    "text": "x",
                    "left": 0,
                    "top": 0,
                    "width": 10,
                    "height": 10,
                },
            )
        assert controller.connection.reset_count == 1

    def test_busy_application_maps_to_connection_error(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes.AddTextbox.side_effect = make_com_error(-2147418111, "zajete")

        with pytest.raises(ComConnectionError) as info:
            controller.dispatch(
                "add_textbox",
                {
                    "slide_index": 1,
                    "text": "x",
                    "left": 0,
                    "top": 0,
                    "width": 10,
                    "height": 10,
                },
            )
        assert "zajety" in info.value.message

    def test_bad_index_maps_to_invalid_reference(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Delete.side_effect = make_com_error(-2147352565, "zly indeks")

        with pytest.raises(InvalidReferenceError):
            controller.dispatch("delete_slide", {"slide_index": 1})


class TestGrouping:
    def test_group_translates_ids_to_shape_indexes(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        second = make_shape(shape_id=6, name="Drugi")
        slides[0].Shapes = com_collection([slides[0].Shapes(1), second])

        controller.dispatch(
            "group_shapes", {"slide_index": 1, "shape_ids": [5, 6], "name": "Karta"}
        )

        slides[0].Shapes.Range.assert_called_once_with([1, 2])
        assert slides[0].Shapes.Range.return_value.Group.return_value.Name == "Karta"

    def test_group_needs_two_shapes(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("group_shapes", {"slide_index": 1, "shape_ids": [5]})

    def test_group_rejects_non_list(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("group_shapes", {"slide_index": 1, "shape_ids": 5})

    def test_ungroup_on_plain_shape_is_invalid_reference(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes(1).Ungroup.side_effect = make_com_error(-2147024809, "nie grupa")

        with pytest.raises(InvalidReferenceError):
            controller.dispatch("ungroup_shapes", {"slide_index": 1, "shape_id": 5})

    def test_align_maps_name_to_constant(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        second = make_shape(shape_id=6)
        slides[0].Shapes = com_collection([slides[0].Shapes(1), second])

        controller.dispatch(
            "align_shapes", {"slide_index": 1, "shape_ids": [5, 6], "align": "middle"}
        )

        slides[0].Shapes.Range.return_value.Align.assert_called_once_with(4, 0)

    def test_align_to_slide_allows_single_shape(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint

        controller.dispatch(
            "align_shapes",
            {
                "slide_index": 1,
                "shape_ids": [5],
                "align": "center",
                "relative_to_slide": True,
            },
        )

        slides[0].Shapes.Range.return_value.Align.assert_called_once_with(1, -1)

    def test_unknown_align_is_rejected(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        second = make_shape(shape_id=6)
        slides[0].Shapes = com_collection([slides[0].Shapes(1), second])

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "align_shapes", {"slide_index": 1, "shape_ids": [5, 6], "align": "srodek"}
            )

    def test_distribute_needs_three_shapes(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        second = make_shape(shape_id=6)
        slides[0].Shapes = com_collection([slides[0].Shapes(1), second])

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "distribute_shapes", {"slide_index": 1, "shape_ids": [5, 6]}
            )

    def test_distribute_vertical(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shapes = [slides[0].Shapes(1), make_shape(shape_id=6), make_shape(shape_id=7)]
        slides[0].Shapes = com_collection(shapes)

        controller.dispatch(
            "distribute_shapes",
            {"slide_index": 1, "shape_ids": [5, 6, 7], "direction": "vertical"},
        )

        slides[0].Shapes.Range.return_value.Distribute.assert_called_once_with(1, 0)


class TestHyperlinks:
    def test_external_url(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        controller.dispatch(
            "add_hyperlink",
            {"slide_index": 1, "shape_id": 5, "url": "https://openai.com"},
        )

        settings = shape.ActionSettings.return_value
        assert shape.ActionSettings.call_args[0][0] == 1  # ppMouseClick
        assert settings.Action == 7  # ppActionHyperlink
        assert settings.Hyperlink.Address == "https://openai.com"

    def test_slide_target_builds_subaddress(self, powerpoint):
        controller, _app, presentation, slides = powerpoint
        target = make_slide(title_text="Skala")
        target.SlideID = 261
        presentation.Slides = com_collection([slides[0], target])

        result = controller.dispatch(
            "add_hyperlink", {"slide_index": 1, "shape_id": 5, "target_slide": 2}
        )

        settings = slides[0].Shapes(1).ActionSettings.return_value
        assert settings.Hyperlink.SubAddress == "261,2,Skala"
        assert result["target_slide"] == 2

    def test_url_and_slide_together_are_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_hyperlink",
                {
                    "slide_index": 1,
                    "shape_id": 5,
                    "url": "https://x.dev",
                    "target_slide": 1,
                },
            )

    def test_neither_url_nor_slide_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_hyperlink", {"slide_index": 1, "shape_id": 5, "tooltip": "hej"}
            )


class TestHeadersFooters:
    def test_footer_text_turns_footer_on(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint

        controller.dispatch(
            "set_headers_footers", {"footer_text": "office-mcp"}
        )

        headers = slides[0].HeadersFooters
        assert headers.Footer.Text == "office-mcp"
        assert headers.Footer.Visible == -1

    def test_explicit_visibility_wins_over_implicit(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint

        controller.dispatch(
            "set_headers_footers",
            {"footer_text": "ukryta", "show_footer": False},
        )

        assert slides[0].HeadersFooters.Footer.Visible == 0

    def test_footer_falls_back_to_master_on_blank_layout(self, powerpoint):
        controller, _app, presentation, slides = powerpoint
        # Uklad "blank" nie ma placeholdera stopki - slajd odrzuca zapis tekstu.
        type(slides[0].HeadersFooters.Footer).Text = property(
            lambda _self: "", lambda _self, _value: (_ for _ in ()).throw(
                make_com_error(-2147352567, "Invalid request")
            )
        )

        result = controller.dispatch(
            "set_headers_footers", {"footer_text": "office-mcp"}
        )

        assert presentation.SlideMaster.HeadersFooters.Footer.Text == "office-mcp"
        assert result["text_on_master"] is True

    def test_slide_numbers_on_single_slide(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint

        result = controller.dispatch(
            "set_headers_footers", {"slide_index": 1, "show_slide_number": True}
        )

        assert slides[0].HeadersFooters.SlideNumber.Visible == -1
        assert result["slides"] == [1]

    def test_no_fields_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_headers_footers", {})


class TestSections:
    def test_add_section_before_slide(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        presentation.SectionProperties.AddBeforeSlide.return_value = 1

        result = controller.dispatch(
            "add_section", {"name": "Wstep", "before_slide": 1}
        )

        presentation.SectionProperties.AddBeforeSlide.assert_called_once_with(1, "Wstep")
        assert result["section_index"] == 1

    def test_add_section_validates_slide(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_section", {"name": "X", "before_slide": 9})

    def test_delete_section_keeps_slides_by_default(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        presentation.SectionProperties.Count = 2
        presentation.SectionProperties.Name.return_value = "Wstep"

        controller.dispatch("delete_section", {"section_index": 1})

        presentation.SectionProperties.Delete.assert_called_once_with(1, 0)

    def test_delete_section_with_slides(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        presentation.SectionProperties.Count = 1

        controller.dispatch(
            "delete_section", {"section_index": 1, "delete_slides": True}
        )

        presentation.SectionProperties.Delete.assert_called_once_with(1, -1)

    def test_list_sections_reads_ranges(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        properties = presentation.SectionProperties
        properties.Count = 1
        properties.Name.return_value = "Historia"
        properties.FirstSlide.return_value = 1
        properties.SlidesCount.return_value = 3

        result = controller.dispatch("list_sections", {})

        assert result["sections"] == [
            {"index": 1, "name": "Historia", "first_slide": 1, "slides": 3}
        ]


class TestSlideshow:
    def test_start_from_given_slide(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint

        controller.dispatch("slideshow", {"command": "start", "slide_index": 1})

        settings = presentation.SlideShowSettings
        assert settings.StartingSlide == 1
        settings.Run.assert_called_once()

    def test_goto_without_slide_index_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("slideshow", {"command": "goto"})

    def test_stop_without_running_show_is_clear_error(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        presentation.SlideShowWindow.View.Exit.side_effect = make_com_error(
            -2147188160, "brak pokazu"
        )

        with pytest.raises(InvalidReferenceError):
            controller.dispatch("slideshow", {"command": "stop"})

    def test_unknown_command_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("slideshow", {"command": "pauza"})


class TestSmartArtAndMedia:
    @staticmethod
    def _layout(name, key, category="process"):
        layout = MagicMock()
        layout.Name = name
        layout.Id = f"urn:microsoft.com/office/officeart/2005/8/layout/{key}"
        layout.Category = category
        return layout

    def test_smartart_layout_matched_by_locale_independent_key(self, powerpoint):
        controller, app, _presentation, slides = powerpoint
        # Polski Office zwraca zlokalizowana nazwe - klucz z URN jest wspolny.
        layout = self._layout("Powtarzany proces łamany", "bProcess3")
        app.SmartArtLayouts = com_collection([layout])
        slides[0].Shapes.AddSmartArt.return_value.SmartArt.AllNodes.Count = 0

        result = controller.dispatch(
            "add_smartart",
            {
                "slide_index": 1,
                "layout": "bProcess3",
                "items": ["Krok 1", {"text": "Szczegol", "level": 2}],
                "left": 0,
                "top": 0,
                "width": 400,
                "height": 200,
            },
        )

        assert slides[0].Shapes.AddSmartArt.call_args[0][0] is layout
        assert result["nodes"] == 2
        # Podwezel powstaje przez AddNode(below) na rodzicu, nie przez Demote()
        root = slides[0].Shapes.AddSmartArt.return_value.SmartArt.AllNodes.Add.return_value
        root.AddNode.assert_called_once_with(5)
        root.Demote.assert_not_called()

    def test_smartart_layout_matched_by_localized_name(self, powerpoint):
        controller, app, _presentation, slides = powerpoint
        layout = self._layout("Podstawowa lista blokowa", "default", "list")
        app.SmartArtLayouts = com_collection([layout])
        slides[0].Shapes.AddSmartArt.return_value.SmartArt.AllNodes.Count = 0

        controller.dispatch(
            "add_smartart",
            {
                "slide_index": 1,
                "layout": "Podstawowa lista blokowa",
                "items": ["A"],
                "left": 0,
                "top": 0,
                "width": 10,
                "height": 10,
            },
        )

        assert slides[0].Shapes.AddSmartArt.call_args[0][0] is layout

    def test_list_smartart_layouts_exposes_key_and_category(self, powerpoint):
        controller, app, _presentation, _slides = powerpoint
        app.SmartArtLayouts = com_collection(
            [
                self._layout("Podstawowa lista blokowa", "default", "list"),
                self._layout("Powtarzany proces łamany", "bProcess3", "process"),
            ]
        )

        result = controller.dispatch("list_smartart_layouts", {"category": "process"})

        assert result["count"] == 1
        assert result["total"] == 2
        assert result["layouts"][0]["key"] == "bProcess3"

    def test_unknown_smartart_layout_is_rejected(self, powerpoint):
        controller, app, _presentation, _slides = powerpoint
        app.SmartArtLayouts = com_collection([self._layout("Proces", "process1")])

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_smartart",
                {
                    "slide_index": 1,
                    "layout": "spirala czasu",
                    "items": ["A"],
                    "left": 0,
                    "top": 0,
                    "width": 10,
                    "height": 10,
                },
            )

    def test_media_falls_back_to_legacy_call(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes.AddMediaObject2.side_effect = AttributeError("brak metody")
        shape = make_shape(shape_id=50)
        shape.MediaType = 3
        slides[0].Shapes.AddMediaObject.return_value = shape

        result = controller.dispatch(
            "add_media",
            {
                "slide_index": 1,
                "media_path": __file__,
                "left": 10,
                "top": 20,
            },
        )

        slides[0].Shapes.AddMediaObject.assert_called_once()
        assert result["media_type"] == "movie"

    def test_missing_media_file_is_document_error(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "add_media",
                {
                    "slide_index": 1,
                    "media_path": r"C:\nie\ma\filmu.mp4",
                    "left": 0,
                    "top": 0,
                },
            )


class TestCopySlideTo:
    def test_same_file_is_rejected(self, powerpoint, tmp_path):
        controller, _app, presentation, _slides = powerpoint
        target = tmp_path / "test.pptx"
        target.write_bytes(b"x")
        presentation.FullName = str(target)

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "copy_slide_to", {"slide_index": 1, "target_path": str(target)}
            )

    def test_unsaved_source_is_rejected(self, powerpoint, tmp_path):
        controller, _app, presentation, _slides = powerpoint
        target = tmp_path / "cel.pptx"
        target.write_bytes(b"x")
        presentation.Path = ""

        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "copy_slide_to", {"slide_index": 1, "target_path": str(target)}
            )

    def test_missing_target_is_document_error(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "copy_slide_to",
                {"slide_index": 1, "target_path": r"C:\nie\ma\celu.pptx"},
            )


class TestTheme:
    def test_set_theme_colors_maps_names_to_scheme_indexes(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        scheme = presentation.SlideMaster.Theme.ThemeColorScheme

        controller.dispatch(
            "set_theme_colors",
            {"colors": {"accent1": "#10A37F", "dark1": "#0B1014", "hyperlink": "blue"}},
        )

        used = [call.args[0] for call in scheme.Colors.call_args_list]
        assert 5 in used  # accent1
        assert 1 in used  # dark1
        assert 11 in used  # hyperlink

    def test_set_theme_colors_rejects_unknown_name(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_theme_colors", {"colors": {"akcent7": "red"}})

    def test_set_theme_colors_rejects_empty_dict(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_theme_colors", {"colors": {}})

    def test_set_theme_fonts_writes_latin_slot(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        fonts = presentation.SlideMaster.Theme.ThemeFontScheme

        controller.dispatch(
            "set_theme_fonts", {"major": "Segoe UI", "minor": "Segoe UI"}
        )

        assert fonts.MajorFont.call_args[0][0] == 1  # msoThemeLatin
        assert fonts.MajorFont.return_value.Name == "Segoe UI"
        assert fonts.MinorFont.return_value.Name == "Segoe UI"

    def test_set_theme_fonts_without_arguments_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_theme_fonts", {})

    def test_master_background_makes_slides_follow_master(self, powerpoint):
        controller, _app, presentation, slides = powerpoint

        result = controller.dispatch(
            "set_master_background", {"color": "#0B1014"}
        )

        fill = presentation.SlideMaster.Background.Fill
        fill.Solid.assert_called_once()
        assert fill.ForeColor.RGB == 0x14100B
        assert slides[0].FollowMasterBackground == -1
        assert result["slides_following_master"] == 1

    def test_master_background_can_skip_slides(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].FollowMasterBackground = 0

        result = controller.dispatch(
            "set_master_background",
            {"color": "#0B1014", "apply_to_slides": False},
        )

        assert slides[0].FollowMasterBackground == 0
        assert result["slides_following_master"] == 0

    def test_master_background_needs_color_or_image(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_master_background", {})


class TestShapeFormat:
    def test_gradient_sets_both_stops(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        controller.dispatch(
            "set_shape_format",
            {
                "slide_index": 1,
                "shape_id": 5,
                "gradient_from": "#10A37F",
                "gradient_to": "#3FE0A0",
                "gradient_style": "diagonal_up",
            },
        )

        shape.Fill.TwoColorGradient.assert_called_once_with(3, 1)
        assert shape.Fill.ForeColor.RGB == 0x7FA310
        assert shape.Fill.BackColor.RGB == 0xA0E03F

    def test_half_a_gradient_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_shape_format",
                {"slide_index": 1, "shape_id": 5, "gradient_from": "#10A37F"},
            )

    def test_transparency_accepts_percent(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        controller.dispatch(
            "set_shape_format",
            {"slide_index": 1, "shape_id": 5, "fill_transparency": 40},
        )

        assert shape.Fill.Transparency == pytest.approx(0.4)

    def test_transparency_out_of_range_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_shape_format",
                {"slide_index": 1, "shape_id": 5, "fill_transparency": 140},
            )

    def test_shadow_details_turn_shadow_on(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        controller.dispatch(
            "set_shape_format",
            {
                "slide_index": 1,
                "shape_id": 5,
                "shadow_blur": 12,
                "shadow_offset_y": 4,
                "shadow_color": "black",
            },
        )

        assert shape.Shadow.Visible == -1
        assert shape.Shadow.Style == 2
        assert shape.Shadow.Blur == 12.0
        assert shape.Shadow.OffsetY == 4.0

    def test_shadow_false_turns_it_off(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        controller.dispatch(
            "set_shape_format", {"slide_index": 1, "shape_id": 5, "shadow": False}
        )

        assert shape.Shadow.Visible == 0

    def test_line_dash_maps_to_constant(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        controller.dispatch(
            "set_shape_format",
            {"slide_index": 1, "shape_id": 5, "line_dash": "round_dot"},
        )

        assert shape.Line.DashStyle == 3

    def test_corner_radius_needs_adjustment_handle(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes(1).Adjustments.Count = 0

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_shape_format",
                {"slide_index": 1, "shape_id": 5, "corner_radius": 0.3},
            )

    def test_corner_radius_uses_setitem(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        adjustments = slides[0].Shapes(1).Adjustments
        adjustments.Count = 1

        controller.dispatch(
            "set_shape_format",
            {"slide_index": 1, "shape_id": 5, "corner_radius": 0.25},
        )

        adjustments.SetItem.assert_called_once_with(1, 0.25)


class TestParagraphFormat:
    def test_applies_spacing_and_anchor(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        result = controller.dispatch(
            "set_paragraph_format",
            {
                "slide_index": 1,
                "shape_id": 5,
                "line_spacing": 1.2,
                "space_after": 8,
                "alignment": "center",
                "vertical_anchor": "middle",
            },
        )

        fmt = shape.TextFrame.TextRange.ParagraphFormat
        assert fmt.SpaceWithin == 1.2
        assert fmt.SpaceAfter == 8.0
        assert fmt.Alignment == 2
        assert shape.TextFrame.VerticalAnchor == 3
        assert result["paragraph"] == "all"

    def test_targets_single_paragraph(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        shape.TextFrame.TextRange.Paragraphs.return_value.Count = 3

        result = controller.dispatch(
            "set_paragraph_format",
            {"slide_index": 1, "shape_id": 5, "paragraph": 2, "space_before": 6},
        )

        assert shape.TextFrame.TextRange.Paragraphs.call_args[0] == (2,)
        assert result["paragraph"] == 2

    def test_paragraph_out_of_range(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes(1).TextFrame.TextRange.Paragraphs.return_value.Count = 2

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_paragraph_format",
                {"slide_index": 1, "shape_id": 5, "paragraph": 9, "space_before": 6},
            )

    def test_autosize_and_margins(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        frame = slides[0].Shapes(1).TextFrame

        controller.dispatch(
            "set_paragraph_format",
            {"slide_index": 1, "shape_id": 5, "autosize": True, "margin": 0},
        )

        assert frame.AutoSize == 1
        assert frame.MarginLeft == 0.0
        assert frame.MarginBottom == 0.0

    def test_no_fields_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_paragraph_format", {"slide_index": 1, "shape_id": 5}
            )

    def test_shape_without_text_frame_is_rejected(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes(1).HasTextFrame = False

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_paragraph_format",
                {"slide_index": 1, "shape_id": 5, "line_spacing": 1.5},
            )


class TestChartFormat:
    def test_non_chart_shape_is_rejected(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        slides[0].Shapes(1).HasChart = False

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "format_chart", {"slide_index": 1, "shape_id": 5, "data_labels": True}
            )

    def test_series_colors_stop_at_series_count(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        shape.HasChart = True
        shape.Chart.SeriesCollection.return_value.Count = 2

        result = controller.dispatch(
            "format_chart",
            {
                "slide_index": 1,
                "shape_id": 5,
                "series_colors": ["#10A37F", "#19C37D", "#3FE0A0"],
            },
        )

        assert result["applied"]["series_colored"] == 2

    def test_transparent_background_hides_chart_area(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        shape.HasChart = True
        shape.Chart.SeriesCollection.return_value.Count = 1

        controller.dispatch(
            "format_chart", {"slide_index": 1, "shape_id": 5, "background": "none"}
        )

        assert shape.Chart.ChartArea.Format.Fill.Visible == 0
        assert shape.Chart.PlotArea.Format.Fill.Visible == 0

    def test_legend_position_maps_to_xl_constant(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        shape.HasChart = True
        shape.Chart.SeriesCollection.return_value.Count = 1

        controller.dispatch(
            "format_chart", {"slide_index": 1, "shape_id": 5, "legend": "bottom"}
        )

        assert shape.Chart.HasLegend == -1
        assert shape.Chart.Legend.Position == -4107

    def test_legend_false_turns_it_off(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        shape.HasChart = True
        shape.Chart.SeriesCollection.return_value.Count = 1

        controller.dispatch(
            "format_chart", {"slide_index": 1, "shape_id": 5, "legend": False}
        )

        assert shape.Chart.HasLegend == 0

    def test_gridlines_target_value_axis(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        shape.HasChart = True
        shape.Chart.SeriesCollection.return_value.Count = 1

        controller.dispatch(
            "format_chart", {"slide_index": 1, "shape_id": 5, "gridlines": False}
        )

        assert shape.Chart.Axes.call_args[0][0] == 2  # xlValue


class TestShapeEditing:
    def test_delete_shape_by_id(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        target = slides[0].Shapes(1)

        result = controller.dispatch(
            "delete_shape", {"slide_index": 1, "shape_id": 5}
        )

        target.Delete.assert_called_once()
        assert result["shape_id"] == 5

    def test_delete_unknown_shape_is_invalid_reference(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("delete_shape", {"slide_index": 1, "shape_id": 999})

    def test_set_shape_position_applies_only_given_fields(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        controller.dispatch(
            "set_shape_position",
            {"slide_index": 1, "shape_id": 5, "left": 100, "rotation": 15},
        )

        assert shape.Left == 100.0
        assert shape.Top == 20.0  # nietkniete
        assert shape.Rotation == 15.0

    def test_set_shape_position_unlocks_aspect_ratio_for_both_sizes(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)
        shape.LockAspectRatio = -1

        controller.dispatch(
            "set_shape_position",
            {"slide_index": 1, "shape_id": 5, "width": 400, "height": 120},
        )

        assert shape.Width == 400.0
        assert shape.Height == 120.0
        assert shape.LockAspectRatio == -1  # blokada przywrocona

    def test_set_shape_position_without_any_field_is_rejected(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_shape_position", {"slide_index": 1, "shape_id": 5})

    def test_set_shape_order_maps_names_to_mso_constants(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = slides[0].Shapes(1)

        for order, expected in (
            ("front", 0),
            ("back", 1),
            ("forward", 2),
            ("backward", 3),
        ):
            controller.dispatch(
                "set_shape_order", {"slide_index": 1, "shape_id": 5, "order": order}
            )
            assert shape.ZOrder.call_args[0][0] == expected

    def test_unknown_order_is_invalid_reference(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_shape_order", {"slide_index": 1, "shape_id": 5, "order": "wyzej"}
            )


class TestExport:
    def test_export_slide_derives_height_from_slide_ratio(self, powerpoint, tmp_path):
        controller, _app, _presentation, slides = powerpoint
        target = tmp_path / "slajd.png"

        result = controller.dispatch(
            "export_slide", {"slide_index": 1, "path": str(target)}
        )

        slides[0].Export.assert_called_once_with(str(target), "PNG", 1920, 1080)
        assert (result["width"], result["height"]) == (1920, 1080)
        assert result["format"] == "PNG"

    def test_export_slide_honours_explicit_width(self, powerpoint, tmp_path):
        controller, _app, _presentation, slides = powerpoint
        target = tmp_path / "slajd.jpg"

        controller.dispatch(
            "export_slide", {"slide_index": 1, "path": str(target), "width": 800}
        )

        slides[0].Export.assert_called_once_with(str(target), "JPG", 800, 450)

    def test_export_slide_rejects_unknown_extension(self, powerpoint, tmp_path):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "export_slide",
                {"slide_index": 1, "path": str(tmp_path / "slajd.svg")},
            )

    def test_export_pdf_saves_copy_without_repointing_presentation(
        self, powerpoint, tmp_path
    ):
        controller, _app, presentation, _slides = powerpoint
        target = tmp_path / "deck.pdf"

        result = controller.dispatch("export_pdf", {"path": str(target)})

        presentation.SaveCopyAs.assert_called_once_with(str(target), 32, -1)
        presentation.SaveAs.assert_not_called()
        assert result["embed_fonts"] is True

    def test_export_pdf_without_embedded_fonts(self, powerpoint, tmp_path):
        controller, _app, presentation, _slides = powerpoint

        controller.dispatch(
            "export_pdf", {"path": str(tmp_path / "d.pdf"), "embed_fonts": False}
        )

        assert presentation.SaveCopyAs.call_args[0][2] == 0

    def test_export_pdf_to_missing_directory_is_document_error(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch(
                "export_pdf", {"path": r"C:\nie\ma\takiego\katalogu\deck.pdf"}
            )


class TestAnimations:
    def test_add_animation_translates_names_to_constants(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        sequence = slides[0].TimeLine.MainSequence
        sequence.Count = 1

        result = controller.dispatch(
            "add_animation",
            {
                "slide_index": 1,
                "shape_id": 5,
                "effect": "rise_up",
                "trigger": "with_previous",
                "level": "by_paragraph",
                "duration": 0.7,
                "delay": 0.2,
            },
        )

        kwargs = sequence.AddEffect.call_args.kwargs
        assert kwargs["effectId"] == 34  # msoAnimEffectRiseUp
        assert kwargs["trigger"] == 2  # msoAnimTriggerWithPrevious
        assert kwargs["Level"] == 1  # msoAnimateTextByAllLevels
        assert kwargs["Shape"].Id == 5

        timing = sequence.AddEffect.return_value.Timing
        assert timing.Duration == 0.7
        assert timing.TriggerDelayTime == 0.2
        assert result["shape_id"] == 5
        assert result["sequence_index"] == 1

    def test_add_animation_accepts_title_shortcut(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        sequence = slides[0].TimeLine.MainSequence
        sequence.Count = 1

        controller.dispatch("add_animation", {"slide_index": 1, "shape_id": "title"})

        assert sequence.AddEffect.call_args.kwargs["Shape"].Name == "Title 1"

    def test_exit_effect_sets_exit_flag(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        sequence = slides[0].TimeLine.MainSequence
        sequence.Count = 1

        controller.dispatch(
            "add_animation",
            {"slide_index": 1, "shape_id": 5, "effect": "fade", "exit_effect": True},
        )

        assert sequence.AddEffect.return_value.Exit == -1

    def test_unknown_effect_is_invalid_reference(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_animation",
                {"slide_index": 1, "shape_id": 5, "effect": "teleportacja"},
            )

    def test_list_animations_reads_sequence(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        effect = MagicMock()
        effect.EffectType = 10  # msoAnimEffectFade
        effect.Exit = 0
        effect.Shape = make_shape(shape_id=5, name="Tresc")
        effect.Timing.TriggerType = 3
        effect.Timing.Duration = 0.5
        effect.Timing.TriggerDelayTime = 0.0
        slides[0].TimeLine.MainSequence = com_collection([effect])
        slides[0].SlideShowTransition.EntryEffect = 3849  # ppEffectFadeSmoothly

        result = controller.dispatch("list_animations", {"slide_index": 1})

        assert result["count"] == 1
        assert result["effects"][0]["effect"] == "fade"
        assert result["effects"][0]["trigger"] == "after_previous"
        assert result["effects"][0]["exit_effect"] is False
        assert result["transition"] == "fade_smoothly"

    def test_set_transition_applies_to_all_slides_by_default(self, powerpoint):
        controller, _app, presentation, _slides = powerpoint
        extra = make_slide()
        presentation.Slides = com_collection([_slides[0], extra])

        result = controller.dispatch(
            "set_transition", {"effect": "push_left", "duration": 0.8}
        )

        assert result["slides"] == [1, 2]
        for slide in (_slides[0], extra):
            assert slide.SlideShowTransition.EntryEffect == 3853  # ppEffectPushLeft
            assert slide.SlideShowTransition.Duration == 0.8
            assert slide.SlideShowTransition.AdvanceOnTime == 0

    def test_set_transition_advance_after_enables_timer(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint

        controller.dispatch(
            "set_transition",
            {"effect": "fade", "slide_index": 1, "advance_after": 5},
        )

        assert slides[0].SlideShowTransition.AdvanceOnTime == -1
        assert slides[0].SlideShowTransition.AdvanceTime == 5.0

    def test_set_transition_rejects_slide_out_of_range(self, powerpoint):
        controller, *_ = powerpoint
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("set_transition", {"effect": "fade", "slide_index": 9})


class TestParagraphText:
    def test_newlines_become_paragraph_separators(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=40)
        slides[0].Shapes.AddTextbox.return_value = shape

        controller.dispatch(
            "add_textbox",
            {
                "slide_index": 1,
                "text": "Pierwszy\nDrugi\r\nTrzeci",
                "left": 0,
                "top": 0,
                "width": 100,
                "height": 50,
            },
        )

        # \n to w COM miekki lamacz wiersza - akapity rozdziela dopiero \r
        assert shape.TextFrame.TextRange.Text == "Pierwszy\rDrugi\rTrzeci"

    def test_shape_text_is_normalized_too(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=41)
        slides[0].Shapes.AddShape.return_value = shape

        controller.dispatch(
            "add_shape",
            {
                "slide_index": 1,
                "shape_type": "rectangle",
                "left": 0,
                "top": 0,
                "width": 10,
                "height": 10,
                "text": "Gora\nDol",
            },
        )

        assert shape.TextFrame.TextRange.Text == "Gora\rDol"

    def test_plain_text_is_untouched(self, powerpoint):
        controller, _app, _presentation, slides = powerpoint
        shape = make_shape(shape_id=42)
        slides[0].Shapes.AddTextbox.return_value = shape

        controller.dispatch(
            "add_textbox",
            {
                "slide_index": 1,
                "text": "Bez lamania",
                "left": 0,
                "top": 0,
                "width": 100,
                "height": 50,
            },
        )

        assert shape.TextFrame.TextRange.Text == "Bez lamania"


class TestSeriesNormalization:
    def test_dict_input(self):
        assert _normalize_series({"A": [1, 2]}, 2) == [("A", [1, 2])]

    def test_list_of_dicts(self):
        assert _normalize_series([{"name": "A", "values": [1]}], 1) == [("A", [1])]

    def test_bare_lists_get_default_names(self):
        assert _normalize_series([[1, 2]], 2) == [("Seria 1", [1, 2])]

    def test_short_series_are_padded(self):
        assert _normalize_series({"A": [1]}, 3) == [("A", [1, None, None])]

    def test_long_series_are_trimmed(self):
        assert _normalize_series({"A": [1, 2, 3]}, 2) == [("A", [1, 2])]

    def test_invalid_type(self):
        with pytest.raises(InvalidReferenceError):
            _normalize_series("dane", 2)
