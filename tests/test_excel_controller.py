from unittest.mock import MagicMock

import pytest

from bridge.controllers.excel import (
    ExcelController,
    _as_formula,
    _column_number,
    _row_span,
)
from bridge.utils.com_helpers import com_address
from bridge.utils.errors import (
    ComConnectionError,
    DocumentNotFoundError,
    InvalidReferenceError,
    UnsupportedOperationError,
)
from tests.conftest import FakeConnection, com_collection, make_com_error


def make_worksheet(name="Arkusz1", values=((1, 2), (3, 4))):
    worksheet = MagicMock()
    worksheet.Name = name
    worksheet.Index = 1
    ranges: dict[str, MagicMock] = {}

    def get_range(reference, *_args):
        if reference not in ranges:
            target = MagicMock()
            target.Address.return_value = f"${reference}"
            target.Value = values
            target.Row = 1
            target.Column = 1
            ranges[reference] = target
        return ranges[reference]

    worksheet.Range.side_effect = get_range
    worksheet.ranges = ranges
    worksheet.UsedRange.Address.return_value = "$A$1:$B$2"
    worksheet.UsedRange.Value = values
    worksheet.UsedRange.Row = 1
    worksheet.UsedRange.Column = 1
    worksheet.UsedRange.Rows.Count = 2
    worksheet.UsedRange.Columns.Count = 2
    return worksheet


def make_workbook(sheets=None, path=r"C:\dane\budzet.xlsx", name="budzet.xlsx"):
    workbook = MagicMock()
    sheet_list = list(sheets or [make_worksheet()])
    workbook.Worksheets = com_collection(sheet_list)
    workbook.Sheets = workbook.Worksheets
    workbook.Name = name
    workbook.Path = path.rsplit("\\", 1)[0] if path else ""
    workbook.FullName = path or ""
    workbook.Saved = True
    workbook.ActiveSheet = sheet_list[0]
    return workbook


@pytest.fixture
def excel():
    worksheet = make_worksheet("Budzet")
    workbook = make_workbook(sheets=[worksheet])

    app = MagicMock()
    app.Workbooks = com_collection([workbook])
    app.ActiveWorkbook = workbook
    app.Workbooks.Add.return_value = workbook
    app.Workbooks.Open.return_value = workbook

    controller = ExcelController(FakeConnection(app=app, key="excel"))
    return controller, app, workbook, worksheet


class TestSheetResolution:
    def test_lookup_by_name_is_case_insensitive(self, excel):
        controller, *_ = excel
        assert controller.worksheet("budzet").Name == "Budzet"

    def test_lookup_by_index(self, excel):
        controller, *_ = excel
        assert controller.worksheet(1).Name == "Budzet"

    def test_unknown_sheet_lists_available(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError) as info:
            controller.worksheet("Koszty")
        assert "Budzet" in info.value.message

    def test_index_out_of_range(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.worksheet(5)

    def test_no_workbook_open(self, excel):
        controller, app, *_ = excel
        app.Workbooks = com_collection([])
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch("get_workbook_info", {})


class TestFileOperations:
    def test_create_workbook_uses_xlsx_format(self, excel, tmp_path):
        controller, app, workbook, _ = excel
        target = tmp_path / "nowy.xlsx"

        controller.dispatch("create_workbook", {"path": str(target)})

        app.Workbooks.Add.assert_called_once()
        assert workbook.SaveAs.call_args[0] == (str(target), 51)

    def test_create_workbook_respects_extension(self, excel, tmp_path):
        controller, _app, workbook, _ = excel
        controller.dispatch("create_workbook", {"path": str(tmp_path / "dane.csv")})
        assert workbook.SaveAs.call_args[0][1] == 6

    def test_open_workbook_reuses_open_file(self, excel, tmp_path):
        controller, app, workbook, _ = excel
        existing = tmp_path / "budzet.xlsx"
        existing.write_bytes(b"x")
        workbook.FullName = str(existing)

        result = controller.dispatch("open_workbook", {"path": str(existing)})

        assert result["already_open"] is True
        app.Workbooks.Open.assert_not_called()

    def test_open_workbook_missing_file(self, excel, tmp_path):
        controller, *_ = excel
        with pytest.raises(DocumentNotFoundError):
            controller.dispatch("open_workbook", {"path": str(tmp_path / "brak.xlsx")})

    def test_save_as_new_path(self, excel, tmp_path):
        controller, _app, workbook, _ = excel
        controller.dispatch("save", {"path": str(tmp_path / "kopia.xlsx")})
        workbook.SaveAs.assert_called_once()

    def test_save_without_file_requires_path(self, excel):
        controller, _app, workbook, _ = excel
        workbook.Path = ""
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("save", {})

    def test_close_with_save(self, excel):
        controller, _app, workbook, _ = excel
        controller.dispatch("close", {})
        workbook.Save.assert_called_once()
        workbook.Close.assert_called_once_with(SaveChanges=True)


class TestSheets:
    def test_add_sheet_appends_at_end(self, excel):
        controller, _app, workbook, _ = excel
        new_sheet = make_worksheet("Nowy")
        workbook.Worksheets.Add.return_value = new_sheet

        result = controller.dispatch("add_sheet", {"name": "Nowy"})

        assert new_sheet.Name == "Nowy"
        assert result["sheet_count"] == 1
        assert "After" in workbook.Worksheets.Add.call_args.kwargs

    def test_add_sheet_with_index_inserts_before(self, excel):
        controller, _app, workbook, _ = excel
        workbook.Worksheets.Add.return_value = make_worksheet("Nowy")

        controller.dispatch("add_sheet", {"name": "Nowy", "index": 1})

        assert "Before" in workbook.Worksheets.Add.call_args.kwargs

    def test_add_sheet_rejects_duplicate_name(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("add_sheet", {"name": "budzet"})

    def test_delete_last_sheet_is_rejected(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("delete_sheet", {"name": "Budzet"})

    def test_delete_sheet(self, excel):
        controller, _app, workbook, worksheet = excel
        second = make_worksheet("Koszty")
        workbook.Worksheets = com_collection([worksheet, second])

        controller.dispatch("delete_sheet", {"name": "Koszty"})

        second.Delete.assert_called_once()

    def test_rename_sheet(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch("rename_sheet", {"old_name": "Budzet", "new_name": "Plan"})
        assert worksheet.Name == "Plan"

    def test_workbook_info_lists_sheets(self, excel):
        controller, *_ = excel
        info = controller.dispatch("get_workbook_info", {})

        assert info["sheets"][0]["name"] == "Budzet"
        assert info["sheets"][0]["used_range"] == "$A$1:$B$2"
        assert info["active_sheet"] == "Budzet"


class TestReadingData:
    def test_get_range_values_returns_matrix(self, excel):
        controller, *_ = excel
        result = controller.dispatch(
            "get_range_values", {"sheet": "Budzet", "range_ref": "A1:B2"}
        )

        assert result["values"] == [[1, 2], [3, 4]]
        assert result["rows"] == 2
        assert result["columns"] == 2

    def test_get_range_values_single_cell(self, excel):
        controller, _app, _workbook, worksheet = excel
        worksheet.Range("A1").Value = 42

        result = controller.dispatch(
            "get_range_values", {"sheet": "Budzet", "range_ref": "A1"}
        )

        assert result["values"] == [[42]]

    def test_get_range_values_rejects_bad_reference(self, excel):
        controller, _app, _workbook, worksheet = excel
        worksheet.Range.side_effect = make_com_error(-2147352571, "zly zakres")

        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "get_range_values", {"sheet": "Budzet", "range_ref": "ZZ!!"}
            )

    def test_get_used_range(self, excel):
        controller, *_ = excel
        result = controller.dispatch("get_used_range", {"sheet": "Budzet"})

        assert result["range"] == "$A$1:$B$2"
        assert result["values"] == [[1, 2], [3, 4]]


class TestWritingData:
    def test_set_cell(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch(
            "set_cell", {"sheet": "Budzet", "cell_ref": "B2", "value": 1500}
        )
        assert worksheet.Range("B2").Value == 1500

    def test_set_range_writes_whole_block(self, excel):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "set_range",
            {
                "sheet": "Budzet",
                "start_cell": "A1",
                "values_2d": [["Nazwa", "Kwota"], ["Serwer", 1200]],
            },
        )

        assert worksheet.Range("A1:B2").Value == (("Nazwa", "Kwota"), ("Serwer", 1200))
        assert result["rows"] == 2

    def test_set_range_block_starts_at_anchor(self, excel):
        controller, _app, _workbook, worksheet = excel
        anchor = worksheet.Range("C3")
        anchor.Row, anchor.Column = 3, 3

        controller.dispatch(
            "set_range",
            {"sheet": "Budzet", "start_cell": "C3", "values_2d": [[1, 2, 3]]},
        )

        assert worksheet.Range("C3:E3").Value == ((1, 2, 3),)

    def test_set_range_pads_ragged_rows(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "set_range",
            {"sheet": "Budzet", "start_cell": "A1", "values_2d": [["a", "b"], ["c"]]},
        )

        assert worksheet.Range("A1:B2").Value == (("a", "b"), ("c", None))

    def test_set_range_rejects_empty_data(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_range", {"sheet": "Budzet", "start_cell": "A1", "values_2d": []}
            )

    def test_set_formula_adds_equals_sign(self, excel):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "set_formula", {"sheet": "Budzet", "cell_ref": "C1", "formula": "SUM(A1:A10)"}
        )

        assert worksheet.Range("C1").Formula == "=SUM(A1:A10)"
        assert result["formula"] == "=SUM(A1:A10)"

    def test_clear_range_contents_only(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch("clear_range", {"sheet": "Budzet", "range_ref": "A1:B2"})
        worksheet.Range("A1:B2").ClearContents.assert_called_once()

    def test_clear_range_with_formatting(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch(
            "clear_range",
            {"sheet": "Budzet", "range_ref": "A1:B2", "contents_only": False},
        )
        worksheet.Range("A1:B2").Clear.assert_called_once()

    def test_insert_rows_uses_span(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch("insert_rows", {"sheet": "Budzet", "start_row": 2, "count": 3})
        worksheet.Rows.assert_called_once_with("2:4")

    def test_delete_rows(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch("delete_rows", {"sheet": "Budzet", "start_row": 5})
        worksheet.Rows.assert_called_once_with("5:5")

    def test_insert_rows_rejects_zero_count(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "insert_rows", {"sheet": "Budzet", "start_row": 1, "count": 0}
            )

    def test_insert_columns_accepts_letter(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch(
            "insert_columns", {"sheet": "Budzet", "start_col": "C", "count": 2}
        )
        worksheet.Columns.assert_called_once_with("C:D")

    def test_insert_columns_accepts_number(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch("insert_columns", {"sheet": "Budzet", "start_col": 27})
        worksheet.Columns.assert_called_once_with("AA:AA")


class TestFormatting:
    def test_set_cell_format_applies_selected_properties(self, excel):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "set_cell_format",
            {
                "sheet": "Budzet",
                "range_ref": "A1:D1",
                "bold": True,
                "fill_color": "#FF0000",
                "number_format": "# ##0,00 zl",
            },
        )

        target = worksheet.Range("A1:D1")
        assert target.Font.Bold is True
        assert target.Interior.Color == 0x0000FF
        assert target.NumberFormat == "# ##0,00 zl"
        assert set(result["applied"]) == {"bold", "fill_color", "number_format"}

    def test_set_cell_format_rejects_unknown_alignment(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_cell_format",
                {"sheet": "Budzet", "range_ref": "A1", "align": "ukosnie"},
            )

    def test_set_column_width_fixed(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch(
            "set_column_width", {"sheet": "Budzet", "column": "B", "width": 22}
        )
        assert worksheet.Columns("B:B").ColumnWidth == 22.0

    def test_set_column_width_auto(self, excel):
        controller, _app, _workbook, worksheet = excel
        result = controller.dispatch(
            "set_column_width", {"sheet": "Budzet", "column": 2, "width": "auto"}
        )
        worksheet.Columns("B:B").AutoFit.assert_called_once()
        assert result["mode"] == "auto"

    def test_merge_cells_centers_by_default(self, excel):
        controller, _app, _workbook, worksheet = excel
        controller.dispatch("merge_cells", {"sheet": "Budzet", "range_ref": "A1:D1"})

        target = worksheet.Range("A1:D1")
        target.Merge.assert_called_once()
        assert target.HorizontalAlignment == -4108

    def test_freeze_panes_toggles_window(self, excel):
        controller, app, _workbook, worksheet = excel
        controller.dispatch("freeze_panes", {"sheet": "Budzet", "cell_ref": "A2"})

        worksheet.Range("A2").Select.assert_called_once()
        assert app.ActiveWindow.FreezePanes is True


class TestConditionalFormatting:
    def test_cell_value_rule(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "apply_conditional_formatting",
            {
                "sheet": "Budzet",
                "range_ref": "B2:B20",
                "rule_type": "cell_value",
                "params": {"operator": "greater", "formula1": 1000, "fill_color": "red"},
            },
        )

        conditions = worksheet.Range("B2:B20").FormatConditions
        conditions.Add.assert_called_once_with(1, 5, "=1000")
        assert conditions.Add.return_value.Interior.Color == 0x0000FF

    def test_cell_value_between_passes_second_formula(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "apply_conditional_formatting",
            {
                "sheet": "Budzet",
                "range_ref": "B2:B20",
                "rule_type": "cell_value",
                "params": {"operator": "between", "formula1": 10, "formula2": 20},
            },
        )

        worksheet.Range("B2:B20").FormatConditions.Add.assert_called_once_with(
            1, 1, "=10", "=20"
        )

    def test_cell_value_requires_threshold(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "apply_conditional_formatting",
                {
                    "sheet": "Budzet",
                    "range_ref": "B2:B20",
                    "rule_type": "cell_value",
                    "params": {"operator": "greater"},
                },
            )

    def test_expression_rule(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "apply_conditional_formatting",
            {
                "sheet": "Budzet",
                "range_ref": "A2:D20",
                "rule_type": "expression",
                "params": {"formula": "=$D2>1000", "bold": True},
            },
        )

        worksheet.Range("A2:D20").FormatConditions.Add.assert_called_once_with(
            2, None, "=$D2>1000"
        )

    def test_color_scale_rule(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "apply_conditional_formatting",
            {
                "sheet": "Budzet",
                "range_ref": "B2:B20",
                "rule_type": "color_scale",
                "params": {"colors": ["#FF0000", "#00FF00"]},
            },
        )

        conditions = worksheet.Range("B2:B20").FormatConditions
        conditions.AddColorScale.assert_called_once_with(ColorScaleType=2)

    def test_unknown_rule_type(self, excel):
        controller, *_ = excel
        with pytest.raises(UnsupportedOperationError):
            controller.dispatch(
                "apply_conditional_formatting",
                {"sheet": "Budzet", "range_ref": "A1", "rule_type": "magia"},
            )


class TestChartsAndTables:
    def test_add_chart_sets_type_and_source(self, excel):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "add_chart",
            {
                "sheet": "Budzet",
                "chart_type": "column",
                "data_range": "A1:B5",
                "left": 200,
                "top": 20,
                "width": 400,
                "height": 250,
                "title": "Koszty",
            },
        )

        chart_objects = worksheet.ChartObjects.return_value
        chart_objects.Add.assert_called_once_with(200.0, 20.0, 400.0, 250.0)
        chart = chart_objects.Add.return_value.Chart
        assert chart.ChartType == 51
        chart.SetSourceData.assert_called_once_with(worksheet.Range("A1:B5"))
        assert result["chart_type"] == "column"

    def test_add_chart_rejects_unknown_type(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_chart",
                {
                    "sheet": "Budzet",
                    "chart_type": "sloneczny",
                    "data_range": "A1:B5",
                    "left": 0,
                    "top": 0,
                    "width": 10,
                    "height": 10,
                },
            )

    def test_create_table(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "create_table",
            {"sheet": "Budzet", "range_ref": "A1:D20", "table_name": "Wydatki"},
        )

        worksheet.ListObjects.Add.assert_called_once_with(
            1, worksheet.Range("A1:D20"), None, 1
        )
        assert worksheet.ListObjects.Add.return_value.Name == "Wydatki"

    def test_create_table_without_headers(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "create_table",
            {
                "sheet": "Budzet",
                "range_ref": "A1:D20",
                "table_name": "Dane",
                "has_headers": False,
            },
        )

        assert worksheet.ListObjects.Add.call_args[0][3] == 2

    def test_add_pivot_table_sets_field_orientation(self, excel):
        controller, _app, workbook, worksheet = excel
        pivot = workbook.PivotCaches.return_value.Create.return_value.CreatePivotTable.return_value
        fields = {}
        pivot.PivotFields.side_effect = lambda name: fields.setdefault(name, MagicMock())

        result = controller.dispatch(
            "add_pivot_table",
            {
                "sheet": "Budzet",
                "source_range": "A1:D50",
                "dest_cell": "F1",
                "rows": ["Kategoria"],
                "columns": ["Miesiac"],
                "values": [{"field": "Kwota", "function": "sum"}],
            },
        )

        assert fields["Kategoria"].Orientation == 1
        assert fields["Miesiac"].Orientation == 2
        pivot.AddDataField.assert_called_once_with(fields["Kwota"], "Sum - Kwota", -4157)
        assert result["source"] == "'Budzet'!$A1:D50"

    def test_add_pivot_table_rejects_unknown_function(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_pivot_table",
                {
                    "sheet": "Budzet",
                    "source_range": "A1:D50",
                    "dest_cell": "F1",
                    "values": [{"field": "Kwota", "function": "mediana"}],
                },
            )


class TestErrorMapping:
    def test_disconnected_excel(self, excel):
        controller, _app, _workbook, worksheet = excel
        worksheet.Range.side_effect = make_com_error(-2147023174, "brak serwera RPC")

        with pytest.raises(ComConnectionError):
            controller.dispatch("set_cell", {"sheet": "Budzet", "cell_ref": "A1", "value": 1})

    def test_generic_com_error_keeps_description(self, excel):
        controller, _app, _workbook, worksheet = excel
        worksheet.Rows.side_effect = make_com_error(-2147352567, "Nie mozna wstawic wierszy")

        with pytest.raises(Exception) as info:
            controller.dispatch("insert_rows", {"sheet": "Budzet", "start_row": 1})
        assert "Nie mozna wstawic wierszy" in str(info.value)


class TestHelpers:
    def test_row_span(self):
        assert _row_span(3, 2) == (3, 4)

    def test_row_span_rejects_zero(self):
        with pytest.raises(InvalidReferenceError):
            _row_span(0, 1)

    def test_row_span_rejects_text(self):
        with pytest.raises(InvalidReferenceError):
            _row_span("drugi", 1)

    def test_column_number_from_letter(self):
        assert _column_number("AA") == 27

    def test_column_number_from_digit_string(self):
        assert _column_number("5") == 5

    def test_column_number_rejects_zero(self):
        with pytest.raises(InvalidReferenceError):
            _column_number(0)

    def test_as_formula(self):
        assert _as_formula(1000) == "=1000"
        assert _as_formula("=A1") == "=A1"

    def test_com_address_handles_method_style_dispatch(self):
        target = MagicMock()
        target.Address.return_value = "$A$1:$C$5"
        assert com_address(target) == "$A$1:$C$5"

    def test_com_address_handles_property_style_dispatch(self):
        target = MagicMock()
        target.Address = "$A$1:$B$2"
        assert com_address(target) == "$A$1:$B$2"

    def test_com_address_falls_back_when_arguments_rejected(self):
        target = MagicMock()
        target.Address.side_effect = [TypeError(), "$D$4"]
        assert com_address(target) == "$D$4"


class TestColumnAndRowSizing:
    def test_delete_columns_builds_span(self, excel):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "delete_columns", {"sheet": "Budzet", "start_col": "C", "count": 2}
        )

        worksheet.Columns.assert_called_with("C:D")
        worksheet.Columns.return_value.Delete.assert_called_once()
        assert result["deleted_columns"] == 2

    def test_delete_columns_accepts_number(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch("delete_columns", {"sheet": "Budzet", "start_col": 3})

        worksheet.Columns.assert_called_with("C:C")

    def test_set_row_height_in_points(self, excel):
        controller, _app, _workbook, worksheet = excel
        worksheet.Rows.return_value.RowHeight = 30.0

        result = controller.dispatch(
            "set_row_height", {"sheet": "Budzet", "row": 1, "height": 30}
        )

        assert worksheet.Rows.call_args[0][0] == 1
        assert result["height"] == 30.0

    def test_set_row_height_auto_calls_autofit(self, excel):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "set_row_height", {"sheet": "Budzet", "row": 2, "height": "auto"}
        )

        worksheet.Rows.return_value.AutoFit.assert_called_once()
        assert result["height"] == "auto"

    def test_row_zero_is_rejected(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "set_row_height", {"sheet": "Budzet", "row": 0, "height": 20}
            )


class TestFindReplaceAndSort:
    def test_empty_needle_is_rejected(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch("find_replace", {"old_text": "", "new_text": "x"})

    def test_no_match_skips_replace(self, excel):
        controller, _app, _workbook, worksheet = excel
        worksheet.UsedRange.Find.return_value = None

        result = controller.dispatch(
            "find_replace",
            {"old_text": "a", "new_text": "b", "sheet": "Budzet", "whole_cell": True},
        )

        worksheet.UsedRange.Replace.assert_not_called()
        assert result["replaced"] == 0

    def test_sort_maps_order_name(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "sort_range",
            {
                "sheet": "Budzet",
                "range_ref": "A1:C4",
                "sort_by": "B",
                "order": "descending",
            },
        )

        kwargs = worksheet.ranges["A1:C4"].Sort.call_args.kwargs
        assert kwargs["Order1"] == 2  # xlDescending
        assert kwargs["Header"] == 1  # xlYes

    def test_sort_pins_orientation_to_columns(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "sort_range", {"sheet": "Budzet", "range_ref": "A1:C4", "sort_by": "B"}
        )

        # Bez jawnego xlSortColumns Excel uzywa "lepkiej" wartosci z poprzedniego
        # sortowania i potrafi poprzestawiac kolumny zamiast wierszy.
        kwargs = worksheet.ranges["A1:C4"].Sort.call_args.kwargs
        assert kwargs["Orientation"] == 1
        assert kwargs["MatchCase"] is False

    def test_sort_without_headers(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "sort_range",
            {
                "sheet": "Budzet",
                "range_ref": "A1:C4",
                "sort_by": 2,
                "has_headers": False,
            },
        )

        assert worksheet.ranges["A1:C4"].Sort.call_args.kwargs["Header"] == 2  # xlNo

    def test_unknown_order_is_rejected(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "sort_range",
                {
                    "sheet": "Budzet",
                    "range_ref": "A1:B2",
                    "sort_by": 1,
                    "order": "rosnaco",
                },
            )


class TestCopyAndValidation:
    def test_copy_all_uses_destination(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "copy_range",
            {"sheet": "Budzet", "range_ref": "A1:B2", "target_cell": "D1"},
        )

        source = worksheet.ranges["A1:B2"]
        assert source.Copy.call_args.kwargs["Destination"] is worksheet.ranges["D1"]

    def test_copy_values_uses_paste_special(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "copy_range",
            {
                "sheet": "Budzet",
                "range_ref": "A1:B2",
                "target_cell": "D1",
                "paste": "values",
            },
        )

        worksheet.ranges["D1"].PasteSpecial.assert_called_once_with(-4163)

    def test_unknown_paste_type_is_rejected(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "copy_range",
                {
                    "sheet": "Budzet",
                    "range_ref": "A1:B2",
                    "target_cell": "D1",
                    "paste": "wszystko",
                },
            )

    def test_validation_list_joins_values(self, excel):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "add_data_validation",
            {"sheet": "Budzet", "range_ref": "E1:E9", "values": ["tak", "nie"]},
        )

        args = worksheet.ranges["E1:E9"].Validation.Add.call_args[0]
        assert args[0] == 3  # xlValidateList
        assert args[3] == "tak,nie"
        assert result["formula1"] == "tak,nie"

    def test_validation_without_values_or_formula_is_rejected(self, excel):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "add_data_validation", {"sheet": "Budzet", "range_ref": "E1"}
            )

    def test_validation_replaces_previous_rule(self, excel):
        controller, _app, _workbook, worksheet = excel

        controller.dispatch(
            "add_data_validation",
            {"sheet": "Budzet", "range_ref": "E1", "values": "tak,nie"},
        )

        worksheet.ranges["E1"].Validation.Delete.assert_called_once()


class TestExcelExport:
    def test_export_pdf_workbook_scope(self, excel, tmp_path):
        controller, _app, workbook, _worksheet = excel
        target = tmp_path / "raport.pdf"

        result = controller.dispatch("export_pdf", {"path": str(target)})

        workbook.ExportAsFixedFormat.assert_called_once_with(0, str(target))
        assert result["scope"] == "workbook"

    def test_export_pdf_sheet_scope(self, excel, tmp_path):
        controller, _app, _workbook, worksheet = excel

        result = controller.dispatch(
            "export_pdf", {"path": str(tmp_path / "a.pdf"), "sheet": "Budzet"}
        )

        worksheet.ExportAsFixedFormat.assert_called_once()
        assert result["scope"] == "sheet"

    def test_export_pdf_range_without_sheet_is_rejected(self, excel, tmp_path):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "export_pdf", {"path": str(tmp_path / "a.pdf"), "range_ref": "A1:B2"}
            )

    def test_range_image_rejects_bad_extension(self, excel, tmp_path):
        controller, *_ = excel
        with pytest.raises(InvalidReferenceError):
            controller.dispatch(
                "export_range_image",
                {
                    "sheet": "Budzet",
                    "range_ref": "A1:B2",
                    "path": str(tmp_path / "a.svg"),
                },
            )

    def test_range_image_removes_helper_chart(self, excel, tmp_path):
        controller, _app, _workbook, worksheet = excel
        target = tmp_path / "zakres.png"
        chart_object = worksheet.ChartObjects.return_value.Add.return_value
        chart_object.Chart.Export.side_effect = lambda path, _fmt: (
            open(path, "wb").write(b"PNG")
        )

        controller.dispatch(
            "export_range_image",
            {"sheet": "Budzet", "range_ref": "A1:B2", "path": str(target)},
        )

        chart_object.Chart.Export.assert_called_once_with(str(target), "PNG")
        chart_object.Delete.assert_called_once()

    def test_range_image_rejects_empty_output(self, excel, tmp_path):
        controller, _app, _workbook, worksheet = excel
        target = tmp_path / "pusty.png"
        chart_object = worksheet.ChartObjects.return_value.Add.return_value
        # Excel potrafi "zapisac" plik zerowej dlugosci i nie zglosic bledu.
        chart_object.Chart.Export.side_effect = lambda path, _fmt: (
            open(path, "wb").close()
        )

        with pytest.raises(UnsupportedOperationError):
            controller.dispatch(
                "export_range_image",
                {"sheet": "Budzet", "range_ref": "A1:B2", "path": str(target)},
            )

        chart_object.Delete.assert_called_once()

    def test_format_chart_without_charts_is_rejected(self, excel):
        controller, _app, _workbook, worksheet = excel
        worksheet.ChartObjects.return_value.Count = 0

        with pytest.raises(InvalidReferenceError):
            controller.dispatch("format_chart", {"sheet": "Budzet"})
