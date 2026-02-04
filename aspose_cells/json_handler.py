"""
Aspose.Cells for Python - JSON Handler Module

This module provides JSON export functionality for workbooks.
It supports exporting worksheet data as a 2D array of values.
"""

import json
from datetime import datetime, date, time
from typing import Optional, Any, List, Dict


class JsonSaveOptions:
    """
    Options for saving JSON files.

    Attributes:
        encoding (str): File encoding. Default is 'utf-8'.
        worksheet_index (int): Index of worksheet to export (-1 for all). Default is -1.
        include_worksheet_name (bool): Include worksheet name in output. Default is True.
        date_format (str): Format string for date values. Default is '%Y-%m-%d'.
        datetime_format (str): Format string for datetime values. Default is '%Y-%m-%d %H:%M:%S'.
        time_format (str): Format string for time values. Default is '%H:%M:%S'.
        empty_cell_value: Value used for empty cells. Default is None (JSON null).
        skip_empty_rows (bool): Skip rows where all cells are empty. Default is False.
        indent (int): JSON indentation level. Default is 2.
        ensure_ascii (bool): Escape non-ASCII characters. Default is False.
    """

    def __init__(self):
        self.encoding = "utf-8"
        self.worksheet_index = -1
        self.include_worksheet_name = True
        self.date_format = "%Y-%m-%d"
        self.datetime_format = "%Y-%m-%d %H:%M:%S"
        self.time_format = "%H:%M:%S"
        self.empty_cell_value = None
        self.skip_empty_rows = False
        self.indent = 2
        self.ensure_ascii = False


class JsonHandler:
    """
    Handles JSON export operations for workbooks.

    Examples:
        >>> wb = Workbook('data.xlsx')
        >>> JsonHandler.save_json(wb, 'output.json')
    """

    @staticmethod
    def save_json(workbook, file_path: str, options: Optional[JsonSaveOptions] = None) -> None:
        """
        Saves a workbook to a JSON file.

        Args:
            workbook: The Workbook object to export.
            file_path (str): Path where the JSON file should be saved.
            options (JsonSaveOptions, optional): Export options. Uses defaults if None.
        """
        if options is None:
            options = JsonSaveOptions()

        data = JsonHandler.save_json_to_dict(workbook, options)

        with open(file_path, "w", encoding=options.encoding) as f:
            json.dump(
                data,
                f,
                indent=options.indent,
                ensure_ascii=options.ensure_ascii,
            )

    @staticmethod
    def save_json_to_dict(workbook, options: Optional[JsonSaveOptions] = None) -> Dict[str, Any]:
        """
        Converts a workbook to a JSON-serializable dictionary.

        Args:
            workbook: The Workbook object to export.
            options (JsonSaveOptions, optional): Export options. Uses defaults if None.

        Returns:
            dict: JSON-serializable data.
        """
        if options is None:
            options = JsonSaveOptions()

        if options.worksheet_index == -1:
            worksheets = [(i, ws) for i, ws in enumerate(workbook.worksheets)]
        else:
            if options.worksheet_index >= len(workbook.worksheets):
                raise IndexError(f"Worksheet index {options.worksheet_index} out of range")
            worksheets = [(options.worksheet_index, workbook.worksheets[options.worksheet_index])]

        sheets_data: List[Dict[str, Any]] = []
        for _, worksheet in worksheets:
            sheet_entry: Dict[str, Any] = {}
            if options.include_worksheet_name:
                sheet_entry["name"] = worksheet.name
            sheet_entry["data"] = JsonHandler._get_worksheet_data(worksheet, options)
            sheets_data.append(sheet_entry)

        return {"worksheets": sheets_data}

    @staticmethod
    def _get_worksheet_data(worksheet, options: JsonSaveOptions) -> List[List[Any]]:
        """
        Extracts all cell data from a worksheet as a 2D list.
        """
        from .cells import Cells

        cells_dict = worksheet.cells._cells
        if not cells_dict:
            return []

        max_row = 0
        max_col = 0
        for ref in cells_dict.keys():
            row, col = Cells.coordinate_from_string(ref)
            if row > max_row:
                max_row = row
            if col > max_col:
                max_col = col

        if max_row == 0 or max_col == 0:
            return []

        rows_data: List[List[Any]] = []
        for row_idx in range(1, max_row + 1):
            row_data: List[Any] = []
            for col_idx in range(1, max_col + 1):
                ref = Cells.coordinate_to_string(row_idx, col_idx)
                cell = cells_dict.get(ref)
                value = cell.value if cell else None
                row_data.append(JsonHandler._format_value(value, options))

            if options.skip_empty_rows and JsonHandler._is_empty_row(row_data, options):
                continue

            rows_data.append(row_data)

        return rows_data

    @staticmethod
    def _is_empty_row(row: List[Any], options: JsonSaveOptions) -> bool:
        """
        Check if a row is empty based on the configured empty_cell_value.
        """
        empty_value = options.empty_cell_value
        for val in row:
            if val is None and empty_value is None:
                continue
            if val == empty_value:
                continue
            if val == "":
                continue
            return False
        return True

    @staticmethod
    def _format_value(value: Any, options: JsonSaveOptions) -> Any:
        """
        Formats a cell value for JSON output.
        """
        if value is None:
            return options.empty_cell_value

        if isinstance(value, datetime):
            return value.strftime(options.datetime_format)
        if isinstance(value, date):
            return value.strftime(options.date_format)
        if isinstance(value, time):
            return value.strftime(options.time_format)
        if isinstance(value, (int, float, bool, str)):
            return value
        return str(value)


def save_workbook_as_json(workbook, file_path: str, options: Optional[JsonSaveOptions] = None) -> None:
    """
    Convenience function to save a Workbook to a JSON file.
    """
    JsonHandler.save_json(workbook, file_path, options)
