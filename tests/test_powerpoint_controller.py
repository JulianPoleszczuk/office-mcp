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
