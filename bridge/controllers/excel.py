"""Excel controller - data, formulas, formatting and charts over COM.

A sheet can be given by name (``"Budget"``) or by number (``1``). Ranges use
A1 notation (``"A1:D10"``), exactly as in the Excel interface.
"""

from __future__ import annotations

import os
import time
from typing import Any

from bridge.controllers.base import BaseController, action, is_connection_error
from bridge.utils.com_helpers import (
    CHART_TYPES,
    XL_COMPARISON_OPERATORS,
    XL_PASTE_TYPES,
    XL_SAVE_FORMATS,
    XL_SORT_ORDERS,
    XL_VALIDATION_ALERTS,
    XL_VALIDATION_TYPES,
    apply_chart_format,
    column_index,
    column_letter,
    com_address,
    com_error,
    from_com_matrix,
    lookup_constant,
    parse_color,
    save_format_for,
    to_com_matrix,
    to_matrix,
    to_python,
)
from bridge.utils.errors import (
    DocumentNotFoundError,
    InvalidReferenceError,
    UnsupportedOperationError,
)

XL_CELL_VALUE = 1
XL_EXPRESSION = 2
XL_WHOLE = 1
XL_PART = 2
XL_BETWEEN = 1
XL_TYPE_PDF = 0
XL_SCREEN = 1
XL_BITMAP = 2
XL_SORT_COLUMNS = 1
XL_TEXT_STRING = 9
XL_CONTAINS = 0
XL_SRC_RANGE = 1
XL_YES = 1
XL_NO = 2
XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_DATA_FIELD = 4

PIVOT_FUNCTIONS = {
    "sum": -4157,
    "count": -4112,
    "average": -4106,
    "max": -4136,
    "min": -4139,
    "product": -4149,
    "count_numbers": -4113,
    "std_dev": -4155,
}

HORIZONTAL_ALIGNMENTS = {
    "left": -4131,
    "center": -4108,
    "right": -4152,
    "justify": -4130,
    "general": 1,
}


class ExcelController(BaseController):
    """``xl_*`` actions - operations on a live Excel instance."""

    APP_KEY = "excel"
    DISPLAY_NAME = "Excel"
    ALERTS_OFF = False

    def workbook(self) -> Any:
        """The active workbook, or a clear error when nothing is open."""
        app = self.app
        if app.Workbooks.Count == 0:
            raise DocumentNotFoundError(
                "No workbook open - use xl_create_workbook or xl_open_workbook"
            )
        try:
            return app.ActiveWorkbook
        except com_error:
            return app.Workbooks(app.Workbooks.Count)

    def worksheet(self, sheet: Any) -> Any:
        """Sheet by name (case-insensitive) or by number."""
        workbook = self.workbook()
        sheets = workbook.Worksheets
        names = [str(sheets(index).Name) for index in range(1, sheets.Count + 1)]

        if isinstance(sheet, int) or (isinstance(sheet, str) and sheet.isdigit()):
            index = self.require_index(sheet, len(names), "sheet")
            return sheets(index)

        wanted = str(sheet).strip().lower()
        for position, name in enumerate(names, start=1):
            if name.lower() == wanted:
                return sheets(position)

        raise InvalidReferenceError(
            f"Sheet '{sheet}' does not exist. Available: {', '.join(names) or 'none'}"
        )

    def range_of(self, worksheet: Any, reference: str) -> Any:
        """An A1 range, with a clear error on a bad address."""
        if not reference or not isinstance(reference, str):
            raise InvalidReferenceError("Range address must be a string, e.g. 'A1:D10'")
        try:
            return worksheet.Range(reference)
        except com_error as exc:
            if is_connection_error(exc):
                raise
            raise InvalidReferenceError(
                f"Invalid range '{reference}' in sheet {worksheet.Name}"
            ) from exc

    def _block_address(self, anchor: Any, rows: int, columns: int) -> str:
        """A1 address of a block of the given size, measured from an anchor cell.

        We deliberately avoid ``Range.Resize`` - under late-bound COM,
        ``Resize(5, 3)`` is sometimes read as the default ``Item`` property and
        returns a single cell instead of a block.
        """
        first_row = int(anchor.Row)
        first_column = int(anchor.Column)
        last_row = first_row + max(1, int(rows)) - 1
        last_column = first_column + max(1, int(columns)) - 1

        return (
            f"{column_letter(first_column)}{first_row}:"
            f"{column_letter(last_column)}{last_row}"
        )

    def _workbook_summary(self, workbook: Any) -> dict[str, Any]:
        return {
            "name": to_python(workbook.Name),
            "path": to_python(workbook.FullName) if workbook.Path else None,
            "sheet_count": int(workbook.Worksheets.Count),
            "saved": bool(workbook.Saved),
        }

    def _activate(self, worksheet: Any) -> None:
        """Switches the view to the sheet so the user sees changes live."""
        try:
            worksheet.Activate()
        except com_error:
            pass

    @action("create_workbook")
    def create_workbook(self, path: str) -> dict[str, Any]:
        """Creates a new workbook and saves it straight to the given path."""
        target = self.resolve_target_path(path)
        workbook = self.app.Workbooks.Add()

        with self.alerts_suppressed():
            workbook.SaveAs(
                target, save_format_for(target, XL_SAVE_FORMATS, XL_SAVE_FORMATS[".xlsx"])
            )
        return self._workbook_summary(workbook)

    @action("open_workbook")
    def open_workbook(self, path: str) -> dict[str, Any]:
        """Opens the file, or activates it if it is already open."""
        target = self.resolve_existing_path(path)
        app = self.app

        for index in range(1, app.Workbooks.Count + 1):
            workbook = app.Workbooks(index)
            if os.path.normcase(str(workbook.FullName)) == os.path.normcase(target):
                try:
                    workbook.Activate()
                except com_error:
                    pass
                return {**self._workbook_summary(workbook), "already_open": True}

        workbook = app.Workbooks.Open(target)
        return {**self._workbook_summary(workbook), "already_open": False}

    @action("save")
    def save(self, path: str | None = None) -> dict[str, Any]:
        """Saves the workbook, or saves it as a new file."""
        workbook = self.workbook()

        if path:
            target = self.resolve_target_path(path)
            with self.alerts_suppressed():
                workbook.SaveAs(
                    target,
                    save_format_for(target, XL_SAVE_FORMATS, XL_SAVE_FORMATS[".xlsx"]),
                )
        elif not workbook.Path:
            raise InvalidReferenceError(
                "The workbook has no file yet - pass the path parameter"
            )
        else:
            workbook.Save()

        return self._workbook_summary(workbook)

    @action("close")
    def close(self, save: bool = True) -> dict[str, Any]:
        """Closes the workbook, optionally saving changes."""
        workbook = self.workbook()
        name = str(workbook.Name)

        if save:
            if not workbook.Path:
                raise InvalidReferenceError(
                    "The workbook was never saved - run xl_save with a path first"
                )
            workbook.Save()

        with self.alerts_suppressed():
            workbook.Close(SaveChanges=bool(save))

        return {"closed": name, "saved": bool(save)}

    @action("add_sheet")
    def add_sheet(self, name: str, index: int | None = None) -> dict[str, Any]:
        """Adds a sheet with the given name; ``index`` sets its position."""
        workbook = self.workbook()
        sheets = workbook.Worksheets
        existing = [str(sheets(i).Name).lower() for i in range(1, sheets.Count + 1)]

        if str(name).lower() in existing:
            raise InvalidReferenceError(f"A sheet named '{name}' already exists")

        if index is None:
            worksheet = sheets.Add(After=sheets(sheets.Count))
        else:
            position = max(1, min(int(index), sheets.Count))
            worksheet = sheets.Add(Before=sheets(position))

        worksheet.Name = str(name)
        return {
            "name": str(name),
            "index": int(worksheet.Index),
            "sheet_count": int(workbook.Worksheets.Count),
        }

    @action("delete_sheet")
    def delete_sheet(self, name: str) -> dict[str, Any]:
        """Deletes a sheet (Excel must keep at least one)."""
        workbook = self.workbook()
        if workbook.Worksheets.Count <= 1:
            raise InvalidReferenceError(
                "Cannot delete the last sheet in a workbook"
            )

        worksheet = self.worksheet(name)
        deleted = str(worksheet.Name)

        with self.alerts_suppressed():
            worksheet.Delete()

        return {"deleted": deleted, "sheet_count": int(workbook.Worksheets.Count)}

    @action("rename_sheet")
    def rename_sheet(self, old_name: str, new_name: str) -> dict[str, Any]:
        """Renames a sheet."""
        worksheet = self.worksheet(old_name)
        worksheet.Name = str(new_name)
        return {"old_name": str(old_name), "new_name": str(new_name)}

    @action("get_workbook_info")
    def get_workbook_info(self) -> dict[str, Any]:
        """Workbook metadata: sheet list, active sheet, path."""
        workbook = self.workbook()
        info = self._workbook_summary(workbook)
        sheets = []

        for index in range(1, workbook.Worksheets.Count + 1):
            worksheet = workbook.Worksheets(index)
            entry = {"index": index, "name": to_python(worksheet.Name)}
            try:
                entry["used_range"] = to_python(com_address(worksheet.UsedRange))
                entry["rows"] = int(worksheet.UsedRange.Rows.Count)
                entry["columns"] = int(worksheet.UsedRange.Columns.Count)
            except com_error:
                entry["used_range"] = None
            sheets.append(entry)

        info["sheets"] = sheets
        try:
            info["active_sheet"] = to_python(workbook.ActiveSheet.Name)
        except com_error:
            info["active_sheet"] = None

        return info

    @action("get_range_values")
    def get_range_values(self, sheet: Any, range_ref: str) -> dict[str, Any]:
        """Reads range values as a 2D array."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)
        values = from_com_matrix(target.Value)

        return {
            "sheet": to_python(worksheet.Name),
            "range": to_python(com_address(target)),
            "rows": len(values),
            "columns": len(values[0]) if values else 0,
            "values": values,
        }

    @action("get_used_range")
    def get_used_range(self, sheet: Any) -> dict[str, Any]:
        """Returns the actually filled area of the sheet, with its data."""
        worksheet = self.worksheet(sheet)
        used = worksheet.UsedRange
        values = from_com_matrix(used.Value)

        return {
            "sheet": to_python(worksheet.Name),
            "range": to_python(com_address(used)),
            "first_row": int(used.Row),
            "first_column": int(used.Column),
            "rows": int(used.Rows.Count),
            "columns": int(used.Columns.Count),
            "values": values,
        }

    @action("set_cell")
    def set_cell(self, sheet: Any, cell_ref: str, value: Any) -> dict[str, Any]:
        """Writes a value into a single cell."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, cell_ref)
        target.Value = value
        self._activate(worksheet)

        return {
            "sheet": to_python(worksheet.Name),
            "cell": to_python(com_address(target)),
            "value": to_python(target.Value),
        }

    @action("set_range")
    def set_range(self, sheet: Any, start_cell: str, values_2d: Any) -> dict[str, Any]:
        """Pastes a whole matrix at once - far faster than cell by cell."""
        matrix = to_matrix(values_2d)
        if not matrix or not matrix[0]:
            raise InvalidReferenceError("Brak danych do wklejenia")

        worksheet = self.worksheet(sheet)
        anchor = self.range_of(worksheet, start_cell)
        target = self.range_of(
            worksheet, self._block_address(anchor, len(matrix), len(matrix[0]))
        )
        target.Value = to_com_matrix(matrix)
        self._activate(worksheet)

        return {
            "sheet": to_python(worksheet.Name),
            "range": to_python(com_address(target)),
            "rows": len(matrix),
            "columns": len(matrix[0]),
        }

    @action("set_formula")
    def set_formula(self, sheet: Any, cell_ref: str, formula: str) -> dict[str, Any]:
        """Writes a formula (``=SUM(A1:A10)``) and returns the computed result."""
        text = str(formula).strip()
        if not text.startswith("="):
            text = "=" + text

        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, cell_ref)
        target.Formula = text
        self._activate(worksheet)

        return {
            "sheet": to_python(worksheet.Name),
            "cell": to_python(com_address(target)),
            "formula": text,
            "value": to_python(target.Value),
        }

    @action("clear_range")
    def clear_range(
        self, sheet: Any, range_ref: str, contents_only: bool = True
    ) -> dict[str, Any]:
        """Clears a range - values only by default, optionally formatting too."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)

        if contents_only:
            target.ClearContents()
        else:
            target.Clear()

        return {
            "sheet": to_python(worksheet.Name),
            "range": to_python(com_address(target)),
            "contents_only": bool(contents_only),
        }

    @action("insert_rows")
    def insert_rows(self, sheet: Any, start_row: int, count: int = 1) -> dict[str, Any]:
        """Inserts rows, pushing existing ones down."""
        first, last = _row_span(start_row, count)
        worksheet = self.worksheet(sheet)
        worksheet.Rows(f"{first}:{last}").Insert()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "inserted_rows": last - first + 1}

    @action("delete_rows")
    def delete_rows(self, sheet: Any, start_row: int, count: int = 1) -> dict[str, Any]:
        """Deletes rows, pulling the rest up."""
        first, last = _row_span(start_row, count)
        worksheet = self.worksheet(sheet)
        worksheet.Rows(f"{first}:{last}").Delete()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "deleted_rows": last - first + 1}

    @action("insert_columns")
    def insert_columns(self, sheet: Any, start_col: Any, count: int = 1) -> dict[str, Any]:
        """Inserts columns; ``start_col`` accepts a letter or a number."""
        first = _column_number(start_col)
        amount = max(1, int(count))
        worksheet = self.worksheet(sheet)
        span = f"{column_letter(first)}:{column_letter(first + amount - 1)}"
        worksheet.Columns(span).Insert()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "inserted_columns": amount, "at": span}

    @action("delete_columns")
    def delete_columns(self, sheet: Any, start_col: Any, count: int = 1) -> dict[str, Any]:
        """Deletes columns; ``start_col`` accepts a letter or a number."""
        first = _column_number(start_col)
        amount = max(1, int(count))
        worksheet = self.worksheet(sheet)
        span = f"{column_letter(first)}:{column_letter(first + amount - 1)}"
        worksheet.Columns(span).Delete()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "deleted_columns": amount, "at": span}

    @action("set_row_height")
    def set_row_height(self, sheet: Any, row: Any, height: Any) -> dict[str, Any]:
        """Row height in points; ``height=\"auto\"`` fits it to the content."""
        worksheet = self.worksheet(sheet)
        index = int(row)
        if index < 1:
            raise InvalidReferenceError("Row number must be >= 1")

        target = worksheet.Rows(index)
        if isinstance(height, str) and height.strip().lower() in ("auto", "autofit"):
            target.AutoFit()
            applied = "auto"
        else:
            target.RowHeight = float(height)
            applied = round(float(target.RowHeight), 2)

        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "row": index, "height": applied}

    @action("find_replace")
    def find_replace(
        self,
        old_text: str,
        new_text: str,
        sheet: Any = None,
        range_ref: str | None = None,
        match_case: bool = False,
        whole_cell: bool = False,
    ) -> dict[str, Any]:
        """Replaces text; without ``sheet`` it walks every sheet.

        ``whole_cell=True`` requires the entire cell content to equal the search
        text - otherwise every matching fragment is replaced.
        """
        if not old_text:
            raise InvalidReferenceError("'old_text' cannot be empty")

        workbook = self.workbook()
        if sheet is None:
            sheets = [
                workbook.Worksheets(index)
                for index in range(1, workbook.Worksheets.Count + 1)
            ]
        else:
            sheets = [self.worksheet(sheet)]

        look_at = XL_WHOLE if whole_cell else XL_PART
        replaced_in: list[str] = []
        for worksheet in sheets:
            target = (
                self.range_of(worksheet, range_ref) if range_ref else worksheet.UsedRange
            )
            before = _count_matches(target, old_text, look_at, match_case)
            if not before:
                continue
            target.Replace(
                What=old_text,
                Replacement=new_text,
                LookAt=look_at,
                MatchCase=bool(match_case),
            )
            replaced_in.append(f"{to_python(worksheet.Name)}:{before}")

        total = sum(int(entry.rsplit(":", 1)[1]) for entry in replaced_in)
        return {
            "replaced": total,
            "sheets": [entry.rsplit(":", 1)[0] for entry in replaced_in],
            "whole_cell": bool(whole_cell),
        }

    @action("sort_range")
    def sort_range(
        self,
        sheet: Any,
        range_ref: str,
        sort_by: Any,
        order: str = "ascending",
        has_headers: bool = True,
    ) -> dict[str, Any]:
        """Sorts a range by column ``sort_by`` (letter, number or cell address)."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)

        if isinstance(sort_by, str) and any(char.isdigit() for char in sort_by):
            key = self.range_of(worksheet, sort_by)
        else:
            column = _column_number(sort_by)
            key = worksheet.Cells(int(target.Row), column)

        # Orientation i MatchCase sa "lepkie" - Excel pamieta je z poprzedniego
        # sort in the session. Without an explicit xlSortColumns it can sort
        # left to right and reorder columns instead of rows.
        target.Sort(
            Key1=key,
            Order1=lookup_constant(order, XL_SORT_ORDERS, "order"),
            Header=XL_YES if has_headers else XL_NO,
            Orientation=XL_SORT_COLUMNS,
            MatchCase=False,
        )

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "range": com_address(target),
            "sorted_by": com_address(key),
            "order": order,
            "has_headers": bool(has_headers),
        }

    @action("set_autofilter")
    def set_autofilter(
        self, sheet: Any, range_ref: str | None = None, enable: bool = True
    ) -> dict[str, Any]:
        """Turns AutoFilter on or off; without ``range_ref`` it covers the used range."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref) if range_ref else worksheet.UsedRange

        already = bool(worksheet.AutoFilterMode)
        if enable and not already:
            target.AutoFilter(1)
        elif not enable and already:
            worksheet.AutoFilterMode = False

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "range": com_address(target),
            "enabled": bool(worksheet.AutoFilterMode),
        }

    @action("copy_range")
    def copy_range(
        self,
        sheet: Any,
        range_ref: str,
        target_cell: str,
        target_sheet: Any = None,
        paste: str = "all",
    ) -> dict[str, Any]:
        """Copies a range; ``paste`` is ``all``, ``values`` or ``formats``."""
        source_sheet = self.worksheet(sheet)
        source = self.range_of(source_sheet, range_ref)
        destination_sheet = (
            self.worksheet(target_sheet) if target_sheet is not None else source_sheet
        )
        destination = self.range_of(destination_sheet, target_cell)
        paste_type = lookup_constant(paste, XL_PASTE_TYPES, "paste")

        if paste_type == XL_PASTE_TYPES["all"]:
            source.Copy(Destination=destination)
        else:
            source.Copy()
            destination.PasteSpecial(paste_type)
            try:
                self.app.CutCopyMode = False
            except com_error:
                pass

        self._activate(destination_sheet)
        return {
            "source": f"{to_python(source_sheet.Name)}!{com_address(source)}",
            "target": f"{to_python(destination_sheet.Name)}!{com_address(destination)}",
            "paste": paste,
            "rows": int(source.Rows.Count),
            "columns": int(source.Columns.Count),
        }

    @action("add_data_validation")
    def add_data_validation(
        self,
        sheet: Any,
        range_ref: str,
        validation_type: str = "list",
        values: Any = None,
        formula: str | None = None,
        formula2: str | None = None,
        operator: str | None = None,
        alert: str = "stop",
        input_message: str | None = None,
        error_message: str | None = None,
    ) -> dict[str, Any]:
        """Data validation - a dropdown list or a range of allowed values.

        For ``validation_type=\"list\"`` just pass ``values`` (a list of entries or
        a range reference). The other types (``whole_number``, ``decimal``,
        ``date``, ``time``, ``text_length``, ``custom``) uzywaja ``formula``,
        ``formula2`` i ``operator``.
        """
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)
        type_constant = lookup_constant(
            validation_type, XL_VALIDATION_TYPES, "validation_type"
        )

        first = formula
        if type_constant == XL_VALIDATION_TYPES["list"] and first is None:
            if isinstance(values, (list, tuple)):
                first = ",".join(str(value) for value in values)
            elif values is not None:
                first = str(values)
        if first is None:
            raise InvalidReferenceError(
                "Pass 'values' (for a list) or 'formula' for the other types"
            )

        operator_constant = (
            lookup_constant(operator, XL_COMPARISON_OPERATORS, "operator")
            if operator
            else XL_BETWEEN
        )

        target.Validation.Delete()
        target.Validation.Add(
            type_constant,
            lookup_constant(alert, XL_VALIDATION_ALERTS, "alert"),
            operator_constant,
            first,
            formula2,
        )

        if input_message:
            target.Validation.InputTitle = ""
            target.Validation.InputMessage = str(input_message)
            target.Validation.ShowInput = True
        if error_message:
            target.Validation.ErrorMessage = str(error_message)
            target.Validation.ShowError = True

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "range": com_address(target),
            "type": validation_type,
            "formula1": first,
            "alert": alert,
        }

    @action("get_cell_formula")
    def get_cell_formula(self, sheet: Any, range_ref: str) -> dict[str, Any]:
        """Returns range formulas (not computed values) along with the results."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)

        formulas: list[list[Any]] = []
        values: list[list[Any]] = []
        for row in range(1, int(target.Rows.Count) + 1):
            formula_row: list[Any] = []
            value_row: list[Any] = []
            for column in range(1, int(target.Columns.Count) + 1):
                cell = target.Cells(row, column)
                formula_row.append(to_python(cell.Formula))
                value_row.append(to_python(cell.Value))
            formulas.append(formula_row)
            values.append(value_row)

        return {
            "sheet": to_python(worksheet.Name),
            "range": com_address(target),
            "formulas": formulas,
            "values": values,
        }

    @action("export_pdf")
    def export_pdf(
        self, path: str, sheet: Any = None, range_ref: str | None = None
    ) -> dict[str, Any]:
        """Exports the workbook, a sheet or a range to PDF.

        W przeciwienstwie do PowerPointa Excel wystawia ``ExportAsFixedFormat``
        in a form pywin32 can call, so there is no need to work around it via
        ``SaveCopyAs``.
        """
        target_path = self.resolve_target_path(path)

        if range_ref is not None:
            if sheet is None:
                raise InvalidReferenceError("'range_ref' requires 'sheet' as well")
            source = self.range_of(self.worksheet(sheet), range_ref)
            scope = "range"
        elif sheet is not None:
            source = self.worksheet(sheet)
            scope = "sheet"
        else:
            source = self.workbook()
            scope = "workbook"

        with self.alerts_suppressed():
            source.ExportAsFixedFormat(XL_TYPE_PDF, target_path)

        return {
            "path": target_path,
            "scope": scope,
            "size_bytes": os.path.getsize(target_path)
            if os.path.isfile(target_path)
            else None,
        }

    @action("export_range_image")
    def export_range_image(
        self, sheet: Any, range_ref: str, path: str
    ) -> dict[str, Any]:
        """Saves a range as a PNG image - a preview for the model.

        Excel cannot export a range to an image directly, so the range goes to
        the clipboard as a bitmap, then onto a temporary chart object, which can
        ``Export``. The chart is removed afterwards.
        """
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)
        target_path = self.resolve_target_path(path)

        extension = os.path.splitext(target_path)[1].lower()
        if extension not in (".png", ".jpg", ".jpeg", ".gif"):
            raise InvalidReferenceError(
                f"Unsupported image extension: {extension or '(none)'}. "
                "Dostepne: .png, .jpg, .jpeg, .gif"
            )

        self._activate(worksheet)

        # CopyPicture gets rejected when the clipboard is still busy after a
        # previous operation (range copy, pivot table).
        # Jedno ponowienie po wyczyszczeniu trybu kopiowania wystarcza.
        for attempt in range(2):
            try:
                target.CopyPicture(XL_SCREEN, XL_BITMAP)
                break
            except com_error as exc:
                if attempt:
                    raise UnsupportedOperationError(
                        "Excel rejected copying the range to the clipboard - close "
                        "any dialog boxes and try again"
                    ) from exc
                try:
                    self.app.CutCopyMode = False
                except com_error:
                    pass
                time.sleep(0.4)

        chart_object = worksheet.ChartObjects().Add(
            0, 0, float(target.Width) + 8, float(target.Height) + 8
        )
        try:
            chart_object.Chart.Paste()
            chart_object.Chart.Export(
                target_path, "JPG" if extension in (".jpg", ".jpeg") else extension[1:].upper()
            )
        finally:
            try:
                chart_object.Delete()
            except com_error:
                pass
            try:
                self.app.CutCopyMode = False
            except com_error:
                pass

        # Chart.Export can report success and leave a zero-length file when the
        # sheet was not active or the bitmap never reached the clipboard.
        if not os.path.isfile(target_path) or os.path.getsize(target_path) == 0:
            raise UnsupportedOperationError(
                "Excel wrote an empty range image - try again after activating "
                "the sheet; with a busy clipboard the export is unreliable"
            )

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "range": com_address(target),
            "path": target_path,
            "size_bytes": os.path.getsize(target_path)
            if os.path.isfile(target_path)
            else None,
        }

    @action("format_chart")
    def format_chart(
        self,
        sheet: Any,
        chart: Any = 1,
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
        """Tunes a chart in the sheet - the counterpart of ``ppt_format_chart``."""
        worksheet = self.worksheet(sheet)
        charts = worksheet.ChartObjects()
        count = int(charts.Count)
        if not count:
            raise InvalidReferenceError(
                f"Sheet {to_python(worksheet.Name)} contains no charts"
            )

        if isinstance(chart, str) and not str(chart).isdigit():
            wanted = str(chart).strip().lower()
            chart_object = None
            for index in range(1, count + 1):
                if str(charts(index).Name).strip().lower() == wanted:
                    chart_object = charts(index)
                    break
            if chart_object is None:
                raise InvalidReferenceError(f"Chart '{chart}' not found")
        else:
            chart_object = charts(self.require_index(chart, count, "chart"))

        applied = apply_chart_format(
            chart_object.Chart,
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

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "chart": to_python(chart_object.Name),
            "applied": applied,
        }

    @action("set_cell_format")
    def set_cell_format(
        self,
        sheet: Any,
        range_ref: str,
        bold: bool | None = None,
        italic: bool | None = None,
        font_size: float | None = None,
        font_color: Any = None,
        fill_color: Any = None,
        number_format: str | None = None,
        align: str | None = None,
        wrap_text: bool | None = None,
    ) -> dict[str, Any]:
        """Formats a range - font, colours, number format, alignment."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)
        applied: dict[str, Any] = {}

        if bold is not None:
            target.Font.Bold = bool(bold)
            applied["bold"] = bool(bold)
        if italic is not None:
            target.Font.Italic = bool(italic)
            applied["italic"] = bool(italic)
        if font_size is not None:
            target.Font.Size = float(font_size)
            applied["font_size"] = float(font_size)
        if font_color is not None:
            target.Font.Color = parse_color(font_color)
            applied["font_color"] = font_color
        if fill_color is not None:
            target.Interior.Color = parse_color(fill_color)
            applied["fill_color"] = fill_color
        if number_format:
            target.NumberFormat = str(number_format)
            applied["number_format"] = str(number_format)
        if align:
            key = str(align).strip().lower()
            if key not in HORIZONTAL_ALIGNMENTS:
                raise InvalidReferenceError(
                    f"Unknown alignment '{align}'. Available: "
                    f"{', '.join(sorted(HORIZONTAL_ALIGNMENTS))}"
                )
            target.HorizontalAlignment = HORIZONTAL_ALIGNMENTS[key]
            applied["align"] = key
        if wrap_text is not None:
            target.WrapText = bool(wrap_text)
            applied["wrap_text"] = bool(wrap_text)

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "range": to_python(com_address(target)),
            "applied": applied,
        }

    @action("set_column_width")
    def set_column_width(self, sheet: Any, column: Any, width: Any) -> dict[str, Any]:
        """Sets column width; ``width=\"auto\"`` fits it to the contents."""
        worksheet = self.worksheet(sheet)
        letter = column_letter(_column_number(column))
        columns = worksheet.Columns(f"{letter}:{letter}")

        if isinstance(width, str) and width.strip().lower() in ("auto", "autofit"):
            columns.AutoFit()
            mode = "auto"
        else:
            columns.ColumnWidth = float(width)
            mode = "fixed"

        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "column": letter, "mode": mode}

    @action("merge_cells")
    def merge_cells(self, sheet: Any, range_ref: str, center: bool = True) -> dict[str, Any]:
        """Merges the cells of a range (centred by default)."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)
        target.Merge()

        if center:
            target.HorizontalAlignment = HORIZONTAL_ALIGNMENTS["center"]

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "range": to_python(com_address(target)),
            "centered": bool(center),
        }

    @action("apply_conditional_formatting")
    def apply_conditional_formatting(
        self,
        sheet: Any,
        range_ref: str,
        rule_type: str,
        params: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Dodaje regule formatowania warunkowego.

        Supported ``rule_type``:

        * ``cell_value`` - ``params``: ``operator`` (``greater``, ``less``,
          ``between``...), ``formula1``, opcjonalnie ``formula2``,
        * ``expression`` - ``params``: ``formula`` (np. ``"=$C2>1000"``),
        * ``text_contains`` - ``params``: ``text``,
        * ``color_scale`` - ``params``: ``colors`` (2 or 3 colours),
        * ``data_bar`` - ``params``: ``color``.

        Result colours are set through ``fill_color``, ``font_color`` and ``bold``.
        """
        settings = dict(params or {})
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)
        kind = str(rule_type).strip().lower()

        if kind == "color_scale":
            colors = settings.get("colors") or ["#F8696B", "#FFEB84", "#63BE7B"]
            scale = target.FormatConditions.AddColorScale(ColorScaleType=len(colors))
            for position, color in enumerate(colors, start=1):
                scale.ColorScaleCriteria(position).FormatColor.Color = parse_color(color)
            self._activate(worksheet)
            return {"rule": kind, "range": to_python(com_address(target))}

        if kind == "data_bar":
            bar = target.FormatConditions.AddDatabar()
            if settings.get("color") is not None:
                bar.BarColor.Color = parse_color(settings["color"])
            self._activate(worksheet)
            return {"rule": kind, "range": to_python(com_address(target))}

        if kind == "cell_value":
            operator = lookup_constant(
                settings.get("operator", "greater"),
                XL_COMPARISON_OPERATORS,
                "operator",
            )
            formula1 = settings.get("formula1", settings.get("value"))
            if formula1 is None:
                raise InvalidReferenceError("Rule cell_value requires the formula1 parameter")

            arguments = [XL_CELL_VALUE, operator, _as_formula(formula1)]
            if settings.get("formula2") is not None:
                arguments.append(_as_formula(settings["formula2"]))
            condition = target.FormatConditions.Add(*arguments)

        elif kind == "expression":
            formula = settings.get("formula")
            if not formula:
                raise InvalidReferenceError("Rule expression requires the formula parameter")
            condition = target.FormatConditions.Add(XL_EXPRESSION, None, str(formula))

        elif kind == "text_contains":
            text = settings.get("text")
            if not text:
                raise InvalidReferenceError("Rule text_contains requires the text parameter")
            condition = target.FormatConditions.Add(
                XL_TEXT_STRING, None, str(text), None, str(text), None, XL_CONTAINS
            )

        else:
            raise UnsupportedOperationError(
                f"Unknown rule type '{rule_type}'. Available: cell_value, expression, "
                "text_contains, color_scale, data_bar"
            )

        if settings.get("fill_color") is not None:
            condition.Interior.Color = parse_color(settings["fill_color"])
        if settings.get("font_color") is not None:
            condition.Font.Color = parse_color(settings["font_color"])
        if settings.get("bold") is not None:
            condition.Font.Bold = bool(settings["bold"])

        self._activate(worksheet)
        return {
            "rule": kind,
            "sheet": to_python(worksheet.Name),
            "range": to_python(com_address(target)),
        }

    @action("freeze_panes")
    def freeze_panes(self, sheet: Any, cell_ref: str) -> dict[str, Any]:
        """Freezes rows and columns above and to the left of the given cell."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, cell_ref)
        self._activate(worksheet)

        window = self.app.ActiveWindow
        window.FreezePanes = False
        target.Select()
        window.FreezePanes = True

        return {
            "sheet": to_python(worksheet.Name),
            "cell": to_python(com_address(target)),
        }

    @action("add_chart")
    def add_chart(
        self,
        sheet: Any,
        chart_type: str,
        data_range: str,
        left: float,
        top: float,
        width: float,
        height: float,
        title: str | None = None,
    ) -> dict[str, Any]:
        """Inserts a chart based on a data range from the same sheet."""
        worksheet = self.worksheet(sheet)
        source = self.range_of(worksheet, data_range)
        chart_constant = lookup_constant(chart_type, CHART_TYPES, "chart_type")

        chart_object = worksheet.ChartObjects().Add(
            float(left), float(top), float(width), float(height)
        )
        chart = chart_object.Chart
        chart.ChartType = chart_constant
        chart.SetSourceData(source)

        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = str(title)

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "chart_name": to_python(chart_object.Name),
            "chart_type": chart_type,
            "data_range": to_python(com_address(source)),
        }

    @action("create_table")
    def create_table(
        self,
        sheet: Any,
        range_ref: str,
        table_name: str,
        has_headers: bool = True,
        style: str = "TableStyleMedium2",
    ) -> dict[str, Any]:
        """Turns a range into a native Excel table (ListObject)."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)

        table = worksheet.ListObjects.Add(
            XL_SRC_RANGE, target, None, XL_YES if has_headers else XL_NO
        )
        table.Name = str(table_name)

        if style:
            try:
                table.TableStyle = str(style)
            except com_error:
                pass

        self._activate(worksheet)
        return {
            "sheet": to_python(worksheet.Name),
            "table_name": str(table_name),
            "range": to_python(com_address(target)),
            "has_headers": bool(has_headers),
        }

    @action("add_pivot_table")
    def add_pivot_table(
        self,
        sheet: Any,
        source_range: str,
        dest_cell: str,
        rows: list[str] | None = None,
        columns: list[str] | None = None,
        values: list[Any] | None = None,
        dest_sheet: Any = None,
        table_name: str = "TabelaPrzestawna1",
    ) -> dict[str, Any]:
        """Builds a pivot table from a source range.

        ``values`` accepts field names (``[\"Amount\"]``) or dictionaries
        ``{"field": "Amount", "function": "average"}``.

        The destination cell is handed to COM as a ``Range`` object - Excel
        odrzuca (E_INVALIDARG) adres tekstowy w notacji A1.
        """
        workbook = self.workbook()
        source_worksheet = self.worksheet(sheet)
        source = self.range_of(source_worksheet, source_range)
        target_worksheet = (
            self.worksheet(dest_sheet) if dest_sheet is not None else source_worksheet
        )
        destination = self.range_of(target_worksheet, dest_cell)

        source_address = f"'{source_worksheet.Name}'!{com_address(source)}"
        destination_address = f"'{target_worksheet.Name}'!{com_address(destination)}"

        cache = workbook.PivotCaches().Create(
            SourceType=XL_DATABASE, SourceData=source_address
        )
        pivot = cache.CreatePivotTable(
            TableDestination=destination, TableName=str(table_name)
        )

        for position, field in enumerate(rows or [], start=1):
            pivot_field = _pivot_field(pivot, field)
            pivot_field.Orientation = XL_ROW_FIELD
            pivot_field.Position = position

        for position, field in enumerate(columns or [], start=1):
            pivot_field = _pivot_field(pivot, field)
            pivot_field.Orientation = XL_COLUMN_FIELD
            pivot_field.Position = position

        added_values = []
        for entry in values or []:
            if isinstance(entry, dict):
                field_name = entry.get("field")
                function_name = str(entry.get("function", "sum")).lower()
            else:
                field_name, function_name = entry, "sum"

            if function_name not in PIVOT_FUNCTIONS:
                raise InvalidReferenceError(
                    f"Unknown aggregate function '{function_name}'. Available: "
                    f"{', '.join(sorted(PIVOT_FUNCTIONS))}"
                )

            pivot.AddDataField(
                _pivot_field(pivot, field_name),
                f"{function_name.capitalize()} - {field_name}",
                PIVOT_FUNCTIONS[function_name],
            )
            added_values.append({"field": field_name, "function": function_name})

        self._activate(target_worksheet)
        return {
            "table_name": str(table_name),
            "source": source_address,
            "destination": destination_address,
            "rows": list(rows or []),
            "columns": list(columns or []),
            "values": added_values,
        }


def _row_span(start_row: Any, count: Any) -> tuple[int, int]:
    """Validates a row range and returns a ``(first, last)`` pair."""
    try:
        first = int(start_row)
        amount = int(count)
    except (TypeError, ValueError) as exc:
        raise InvalidReferenceError("Row number and row count must be numbers") from exc

    if first < 1:
        raise InvalidReferenceError("Row number must be >= 1")
    if amount < 1:
        raise InvalidReferenceError("Row count must be >= 1")

    return first, first + amount - 1


def _column_number(column: Any) -> int:
    """Accepts ``\"C\"`` or ``3`` and returns the column number."""
    if isinstance(column, int):
        if column < 1:
            raise InvalidReferenceError("Column number must be >= 1")
        return column
    text = str(column).strip()
    if text.isdigit():
        return _column_number(int(text))
    return column_index(text)


def _as_formula(value: Any) -> str:
    """Turns a threshold value into a formula FormatConditions accepts."""
    text = str(value)
    return text if text.startswith("=") else f"={text}"


def _pivot_field(pivot: Any, field_name: Any) -> Any:
    """Pivot field with a clear error when the name does not match."""
    if not field_name:
        raise InvalidReferenceError("Pivot field name cannot be empty")
    try:
        return pivot.PivotFields(str(field_name))
    except com_error as exc:
        if is_connection_error(exc):
            raise
        raise InvalidReferenceError(
            f"The pivot table has no field '{field_name}' - check the range headers"
        ) from exc


def _count_matches(target: Any, needle: str, look_at: int, match_case: bool) -> int:
    """Counts cells matching the search text before replacing.

    ``Range.Replace`` only returns ``True``/``False``, so the number of hits
    has to be counted separately via ``Find``/``FindNext``.
    """
    try:
        found = target.Find(
            What=needle, LookAt=look_at, MatchCase=bool(match_case), LookIn=-4163
        )
    except com_error:
        return 0
    if found is None:
        return 0

    first = com_address(found)
    seen = {first}
    while True:
        try:
            found = target.FindNext(found)
        except com_error:
            break
        if found is None:
            break
        address = com_address(found)
        if address in seen:
            break
        seen.add(address)
    return len(seen)


__all__ = ["ExcelController"]

