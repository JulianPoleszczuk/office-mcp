"""Kontroler Excela - dane, formuly, formatowanie i wykresy przez COM.

Arkusz mozna wskazac nazwa (``"Budzet"``) albo numerem (``1``). Zakresy
podaje sie w notacji A1 (``"A1:D10"``), tak jak w interfejsie Excela.
"""

from __future__ import annotations

import os
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
    """Akcje ``xl_*`` - operacje na zywej instancji Excela."""

    APP_KEY = "excel"
    DISPLAY_NAME = "Excel"
    ALERTS_OFF = False

    def workbook(self) -> Any:
        """Aktywny skoroszyt albo czytelny blad, gdy nic nie jest otwarte."""
        app = self.app
        if app.Workbooks.Count == 0:
            raise DocumentNotFoundError(
                "Brak otwartego skoroszytu - uzyj xl_create_workbook albo xl_open_workbook"
            )
        try:
            return app.ActiveWorkbook
        except com_error:
            return app.Workbooks(app.Workbooks.Count)

    def worksheet(self, sheet: Any) -> Any:
        """Arkusz po nazwie (bez rozroznienia wielkosci liter) albo po numerze."""
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
            f"Arkusz '{sheet}' nie istnieje. Dostepne: {', '.join(names) or 'brak'}"
        )

    def range_of(self, worksheet: Any, reference: str) -> Any:
        """Zakres A1 z czytelnym bledem przy zlym adresie."""
        if not reference or not isinstance(reference, str):
            raise InvalidReferenceError("Adres zakresu musi byc tekstem, np. 'A1:D10'")
        try:
            return worksheet.Range(reference)
        except com_error as exc:
            if is_connection_error(exc):
                raise
            raise InvalidReferenceError(
                f"Nieprawidlowy zakres '{reference}' w arkuszu {worksheet.Name}"
            ) from exc

    def _block_address(self, anchor: Any, rows: int, columns: int) -> str:
        """Adres A1 bloku o zadanym rozmiarze, liczony od komorki zakotwiczenia.

        Swiadomie nie uzywamy ``Range.Resize`` - przy pozno wiazanym COM
        ``Resize(5, 3)`` bywa interpretowane jako domyslna wlasciwosc ``Item``
        i zwraca pojedyncza komorke zamiast bloku.
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
        """Przelacza widok na arkusz, zeby uzytkownik widzial zmiany na zywo."""
        try:
            worksheet.Activate()
        except com_error:
            pass

    @action("create_workbook")
    def create_workbook(self, path: str) -> dict[str, Any]:
        """Tworzy nowy skoroszyt i od razu zapisuje go pod wskazana sciezka."""
        target = self.resolve_target_path(path)
        workbook = self.app.Workbooks.Add()

        with self.alerts_suppressed():
            workbook.SaveAs(
                target, save_format_for(target, XL_SAVE_FORMATS, XL_SAVE_FORMATS[".xlsx"])
            )
        return self._workbook_summary(workbook)

    @action("open_workbook")
    def open_workbook(self, path: str) -> dict[str, Any]:
        """Otwiera plik albo aktywuje go, jesli jest juz otwarty."""
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
        """Zapisuje skoroszyt albo zapisuje go jako nowy plik."""
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
                "Skoroszyt nie ma jeszcze pliku - podaj parametr path"
            )
        else:
            workbook.Save()

        return self._workbook_summary(workbook)

    @action("close")
    def close(self, save: bool = True) -> dict[str, Any]:
        """Zamyka skoroszyt, opcjonalnie zapisujac zmiany."""
        workbook = self.workbook()
        name = str(workbook.Name)

        if save:
            if not workbook.Path:
                raise InvalidReferenceError(
                    "Skoroszyt nie byl zapisany - najpierw xl_save z parametrem path"
                )
            workbook.Save()

        with self.alerts_suppressed():
            workbook.Close(SaveChanges=bool(save))

        return {"closed": name, "saved": bool(save)}

    @action("add_sheet")
    def add_sheet(self, name: str, index: int | None = None) -> dict[str, Any]:
        """Dodaje arkusz o podanej nazwie; ``index`` ustawia jego pozycje."""
        workbook = self.workbook()
        sheets = workbook.Worksheets
        existing = [str(sheets(i).Name).lower() for i in range(1, sheets.Count + 1)]

        if str(name).lower() in existing:
            raise InvalidReferenceError(f"Arkusz o nazwie '{name}' juz istnieje")

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
        """Usuwa arkusz (Excel musi zostac z co najmniej jednym)."""
        workbook = self.workbook()
        if workbook.Worksheets.Count <= 1:
            raise InvalidReferenceError(
                "Nie mozna usunac ostatniego arkusza w skoroszycie"
            )

        worksheet = self.worksheet(name)
        deleted = str(worksheet.Name)

        with self.alerts_suppressed():
            worksheet.Delete()

        return {"deleted": deleted, "sheet_count": int(workbook.Worksheets.Count)}

    @action("rename_sheet")
    def rename_sheet(self, old_name: str, new_name: str) -> dict[str, Any]:
        """Zmienia nazwe arkusza."""
        worksheet = self.worksheet(old_name)
        worksheet.Name = str(new_name)
        return {"old_name": str(old_name), "new_name": str(new_name)}

    @action("get_workbook_info")
    def get_workbook_info(self) -> dict[str, Any]:
        """Metadane skoroszytu: lista arkuszy, aktywny arkusz, sciezka."""
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
        """Odczytuje wartosci zakresu jako tablice 2D."""
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
        """Zwraca faktycznie wypelniony obszar arkusza wraz z danymi."""
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
        """Wpisuje wartosc do pojedynczej komorki."""
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
        """Wkleja macierz danych naraz - duzo szybciej niz komorka po komorce."""
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
        """Wpisuje formule (``=SUM(A1:A10)``) i zwraca wyliczony wynik."""
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
        """Czysci zakres - domyslnie same wartosci, opcjonalnie takze formatowanie."""
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
        """Wstawia wiersze, przesuwajac istniejace w dol."""
        first, last = _row_span(start_row, count)
        worksheet = self.worksheet(sheet)
        worksheet.Rows(f"{first}:{last}").Insert()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "inserted_rows": last - first + 1}

    @action("delete_rows")
    def delete_rows(self, sheet: Any, start_row: int, count: int = 1) -> dict[str, Any]:
        """Usuwa wiersze, przesuwajac pozostale w gore."""
        first, last = _row_span(start_row, count)
        worksheet = self.worksheet(sheet)
        worksheet.Rows(f"{first}:{last}").Delete()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "deleted_rows": last - first + 1}

    @action("insert_columns")
    def insert_columns(self, sheet: Any, start_col: Any, count: int = 1) -> dict[str, Any]:
        """Wstawia kolumny; ``start_col`` przyjmuje litere albo numer."""
        first = _column_number(start_col)
        amount = max(1, int(count))
        worksheet = self.worksheet(sheet)
        span = f"{column_letter(first)}:{column_letter(first + amount - 1)}"
        worksheet.Columns(span).Insert()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "inserted_columns": amount, "at": span}

    @action("delete_columns")
    def delete_columns(self, sheet: Any, start_col: Any, count: int = 1) -> dict[str, Any]:
        """Usuwa kolumny; ``start_col`` przyjmuje litere albo numer."""
        first = _column_number(start_col)
        amount = max(1, int(count))
        worksheet = self.worksheet(sheet)
        span = f"{column_letter(first)}:{column_letter(first + amount - 1)}"
        worksheet.Columns(span).Delete()
        self._activate(worksheet)
        return {"sheet": to_python(worksheet.Name), "deleted_columns": amount, "at": span}

    @action("set_row_height")
    def set_row_height(self, sheet: Any, row: Any, height: Any) -> dict[str, Any]:
        """Wysokosc wiersza w punktach; ``height="auto"`` dopasowuje do tresci."""
        worksheet = self.worksheet(sheet)
        index = int(row)
        if index < 1:
            raise InvalidReferenceError("Numer wiersza musi byc >= 1")

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
        """Podmienia tekst; bez ``sheet`` przechodzi przez wszystkie arkusze.

        ``whole_cell=True`` wymaga, zeby cala zawartosc komorki byla rowna
        szukanemu tekstowi - inaczej podmieniany jest kazdy fragment.
        """
        if not old_text:
            raise InvalidReferenceError("'old_text' nie moze byc puste")

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
        """Sortuje zakres po kolumnie ``sort_by`` (litera, numer albo adres komorki)."""
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)

        if isinstance(sort_by, str) and any(char.isdigit() for char in sort_by):
            key = self.range_of(worksheet, sort_by)
        else:
            column = _column_number(sort_by)
            key = worksheet.Cells(int(target.Row), column)

        # Orientation i MatchCase sa "lepkie" - Excel pamieta je z poprzedniego
        # sortowania w sesji. Bez jawnego xlSortColumns potrafi posortowac
        # lewo-prawo i poprzestawiac kolumny zamiast wierszy.
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
        """Wlacza albo wylacza autofiltr; bez ``range_ref`` obejmuje uzyty obszar."""
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
        """Kopiuje zakres; ``paste`` to ``all``, ``values`` albo ``formats``."""
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
        """Sprawdzanie poprawnosci danych - lista rozwijana albo zakres wartosci.

        Dla ``validation_type="list"`` wystarczy ``values`` (lista pozycji albo
        odwolanie do zakresu). Pozostale typy (``whole_number``, ``decimal``,
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
                "Podaj 'values' (dla listy) albo 'formula' dla pozostalych typow"
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
        """Zwraca formuly zakresu (a nie wyliczone wartosci) wraz z wynikami."""
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
        """Eksportuje skoroszyt, arkusz albo zakres do PDF-u.

        W przeciwienstwie do PowerPointa Excel wystawia ``ExportAsFixedFormat``
        w formie wywolywalnej przez pywin32, wiec nie trzeba obchodzic tego
        przez ``SaveCopyAs``.
        """
        target_path = self.resolve_target_path(path)

        if range_ref is not None:
            if sheet is None:
                raise InvalidReferenceError("'range_ref' wymaga podania 'sheet'")
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
        """Zapisuje zakres jako obraz PNG - podglad dla modelu.

        Excel nie ma bezposredniego eksportu zakresu do obrazu, wiec zakres
        trafia do schowka jako bitmapa, potem na tymczasowy obiekt wykresu,
        ktory juz potrafi ``Export``. Wykres jest usuwany na koncu.
        """
        worksheet = self.worksheet(sheet)
        target = self.range_of(worksheet, range_ref)
        target_path = self.resolve_target_path(path)

        extension = os.path.splitext(target_path)[1].lower()
        if extension not in (".png", ".jpg", ".jpeg", ".gif"):
            raise InvalidReferenceError(
                f"Nieobslugiwane rozszerzenie obrazu: {extension or '(brak)'}. "
                "Dostepne: .png, .jpg, .jpeg, .gif"
            )

        target.CopyPicture(XL_SCREEN, XL_BITMAP)
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
    ) -> dict[str, Any]:
        """Dostraja wykres w arkuszu - odpowiednik ``ppt_format_chart``."""
        worksheet = self.worksheet(sheet)
        charts = worksheet.ChartObjects()
        count = int(charts.Count)
        if not count:
            raise InvalidReferenceError(
                f"Arkusz {to_python(worksheet.Name)} nie zawiera wykresow"
            )

        if isinstance(chart, str) and not str(chart).isdigit():
            wanted = str(chart).strip().lower()
            chart_object = None
            for index in range(1, count + 1):
                if str(charts(index).Name).strip().lower() == wanted:
                    chart_object = charts(index)
                    break
            if chart_object is None:
                raise InvalidReferenceError(f"Nie znaleziono wykresu '{chart}'")
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
        """Formatuje zakres - czcionka, kolory, format liczb, wyrownanie."""
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
                    f"Nieznane wyrownanie '{align}'. Dostepne: "
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
        """Ustawia szerokosc kolumny; ``width="auto"`` dopasowuje do zawartosci."""
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
        """Scala komorki zakresu (domyslnie z wysrodkowaniem zawartosci)."""
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

        Obslugiwane ``rule_type``:

        * ``cell_value`` - ``params``: ``operator`` (``greater``, ``less``,
          ``between``...), ``formula1``, opcjonalnie ``formula2``,
        * ``expression`` - ``params``: ``formula`` (np. ``"=$C2>1000"``),
        * ``text_contains`` - ``params``: ``text``,
        * ``color_scale`` - ``params``: ``colors`` (2 lub 3 kolory),
        * ``data_bar`` - ``params``: ``color``.

        Kolory wyniku ustawia sie przez ``fill_color``, ``font_color`` i ``bold``.
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
                raise InvalidReferenceError("Regula cell_value wymaga parametru formula1")

            arguments = [XL_CELL_VALUE, operator, _as_formula(formula1)]
            if settings.get("formula2") is not None:
                arguments.append(_as_formula(settings["formula2"]))
            condition = target.FormatConditions.Add(*arguments)

        elif kind == "expression":
            formula = settings.get("formula")
            if not formula:
                raise InvalidReferenceError("Regula expression wymaga parametru formula")
            condition = target.FormatConditions.Add(XL_EXPRESSION, None, str(formula))

        elif kind == "text_contains":
            text = settings.get("text")
            if not text:
                raise InvalidReferenceError("Regula text_contains wymaga parametru text")
            condition = target.FormatConditions.Add(
                XL_TEXT_STRING, None, str(text), None, str(text), None, XL_CONTAINS
            )

        else:
            raise UnsupportedOperationError(
                f"Nieznany typ reguly '{rule_type}'. Dostepne: cell_value, expression, "
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
        """Zamraza wiersze i kolumny powyzej/na lewo od wskazanej komorki."""
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
        """Wstawia wykres oparty o zakres danych z tego samego arkusza."""
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
        """Zamienia zakres w natywna tabele Excela (ListObject)."""
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
        """Buduje tabele przestawna z zakresu zrodlowego.

        ``values`` przyjmuje nazwy pol (``["Kwota"]``) albo slowniki
        ``{"field": "Kwota", "function": "average"}``.

        Komorke docelowa przekazujemy do COM jako obiekt ``Range`` - Excel
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
                    f"Nieznana funkcja agregujaca '{function_name}'. Dostepne: "
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
    """Waliduje zakres wierszy i zwraca pare ``(pierwszy, ostatni)``."""
    try:
        first = int(start_row)
        amount = int(count)
    except (TypeError, ValueError) as exc:
        raise InvalidReferenceError("Numer wiersza i liczba wierszy musza byc liczbami") from exc

    if first < 1:
        raise InvalidReferenceError("Numer wiersza musi byc >= 1")
    if amount < 1:
        raise InvalidReferenceError("Liczba wierszy musi byc >= 1")

    return first, first + amount - 1


def _column_number(column: Any) -> int:
    """Przyjmuje ``"C"`` albo ``3`` i zwraca numer kolumny."""
    if isinstance(column, int):
        if column < 1:
            raise InvalidReferenceError("Numer kolumny musi byc >= 1")
        return column
    text = str(column).strip()
    if text.isdigit():
        return _column_number(int(text))
    return column_index(text)


def _as_formula(value: Any) -> str:
    """Zamienia wartosc progu na formule akceptowana przez FormatConditions."""
    text = str(value)
    return text if text.startswith("=") else f"={text}"


def _pivot_field(pivot: Any, field_name: Any) -> Any:
    """Pole tabeli przestawnej z czytelnym bledem, gdy nazwa nie pasuje."""
    if not field_name:
        raise InvalidReferenceError("Nazwa pola tabeli przestawnej nie moze byc pusta")
    try:
        return pivot.PivotFields(str(field_name))
    except com_error as exc:
        if is_connection_error(exc):
            raise
        raise InvalidReferenceError(
            f"Tabela przestawna nie ma pola '{field_name}' - sprawdz naglowki zakresu"
        ) from exc


def _count_matches(target: Any, needle: str, look_at: int, match_case: bool) -> int:
    """Liczy komorki pasujace do szukanego tekstu przed podmiana.

    ``Range.Replace`` zwraca tylko ``True``/``False``, wiec liczbe trafien
    trzeba policzyc osobno przez ``Find``/``FindNext``.
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

