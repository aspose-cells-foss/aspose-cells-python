"""
Aspose.Cells for Python - Markdown Handler Module

This module provides Markdown export functionality for workbooks.
It supports exporting worksheet data as Markdown tables with configurable
alignment, formatting, and styling options.

Features:
- Export worksheet data to Markdown table format
- Configurable column alignment (left, center, right)
- Support for multiple worksheets export
- Customizable date/time formatting
- Option to include worksheet names as headers
"""

import io
from datetime import datetime, date, time
from typing import Optional, List, Any, Dict


class MarkdownSaveOptions:
    """
    Options for saving Markdown files.

    Attributes:
        encoding (str): File encoding. Default is 'utf-8'.
        worksheet_index (int): Index of worksheet to export (-1 for all). Default is 0.
        include_worksheet_name (bool): Include worksheet name as header. Default is True.
        header_level (int): Markdown header level for worksheet name (1-6). Default is 2.
        column_alignments (dict): Column alignment overrides {col_index: 'left'|'center'|'right'}.
        default_alignment (str): Default column alignment. Default is 'left'.
        date_format (str): Format string for date values. Default is '%Y-%m-%d'.
        datetime_format (str): Format string for datetime values. Default is '%Y-%m-%d %H:%M:%S'.
        time_format (str): Format string for time values. Default is '%H:%M:%S'.
        empty_cell_placeholder (str): Text for empty cells. Default is '' (empty string).
        escape_pipes (bool): Escape pipe characters in cell values. Default is True.
        first_row_as_header (bool): Treat first row as table header. Default is True.
        include_row_numbers (bool): Include row numbers as first column. Default is False.
        max_column_width (int): Maximum column width (0 for unlimited). Default is 0.
        trim_whitespace (bool): Trim whitespace from cell values. Default is True.
        float_precision (int): Decimal places for floats (-1 for auto/smart rounding). Default is -1.
        skip_empty_rows (bool): Skip rows where all cells are empty. Default is False.
        newline_replacement (str): Replace newlines in cell values with this string. Default is ' '.
        detect_title_rows (bool): Detect single-cell rows and output as headings. Default is False.
        auto_detect_header (bool): Auto-detect actual header row (skip title rows). Default is False.
        compact_format (bool): Use compact format without cell padding. Default is False.
        simple_separators (bool): Use simple '---' separators without alignment colons. Default is False.

    Examples:
        >>> options = MarkdownSaveOptions()
        >>> options.default_alignment = 'center'
        >>> options.include_worksheet_name = True
        >>> MarkdownHandler.save_markdown(workbook, 'output.md', options)
    """

    def __init__(self):
        self.encoding = 'utf-8'
        self.worksheet_index = -1
        self.include_worksheet_name = True
        self.header_level = 2
        self.column_alignments: Dict[int, str] = {}
        self.default_alignment = 'left'
        self.date_format = '%Y-%m-%d'
        self.datetime_format = '%Y-%m-%d %H:%M:%S'
        self.time_format = '%H:%M:%S'
        self.empty_cell_placeholder = ''
        self.escape_pipes = True
        self.first_row_as_header = True
        self.include_row_numbers = False
        self.max_column_width = 0
        self.trim_whitespace = True
        self.float_precision = -1  # -1 means auto (remove trailing zeros), 0+ means fixed decimals
        self.skip_empty_rows = False  # Skip rows that are entirely empty
        self.newline_replacement = ' '  # Replace newlines in cell values with this string
        self.detect_title_rows = False  # Detect single-cell rows and output as headings
        self.auto_detect_header = False  # Auto-detect actual header row (skip title rows)
        self.compact_format = True  # Use compact format without cell padding
        self.simple_separators = False  # Use simple '---' separators without alignment colons


class MarkdownHandler:
    """
    Handles Markdown export operations for workbooks.

    This class provides static methods to export workbook data to Markdown format,
    creating properly formatted Markdown tables.

    Examples:
        # Export workbook to Markdown
        >>> wb = Workbook('data.xlsx')
        >>> MarkdownHandler.save_markdown(wb, 'output.md')

        # Export with custom options
        >>> options = MarkdownSaveOptions()
        >>> options.default_alignment = 'center'
        >>> MarkdownHandler.save_markdown(wb, 'output.md', options)
    """

    @staticmethod
    def save_markdown(workbook, file_path: str, options: Optional[MarkdownSaveOptions] = None) -> None:
        """
        Saves a workbook to a Markdown file.

        Args:
            workbook: The Workbook object to export.
            file_path (str): Path where the Markdown file should be saved.
            options (MarkdownSaveOptions, optional): Export options. Uses defaults if None.

        Raises:
            IndexError: If the specified worksheet index is out of range.
            IOError: If the file cannot be written.

        Examples:
            >>> wb = Workbook('data.xlsx')
            >>> MarkdownHandler.save_markdown(wb, 'output.md')

            >>> options = MarkdownSaveOptions()
            >>> options.include_worksheet_name = False
            >>> MarkdownHandler.save_markdown(wb, 'output.md', options)
        """
        if options is None:
            options = MarkdownSaveOptions()

        content = MarkdownHandler.save_markdown_to_string(workbook, options)

        with open(file_path, 'w', encoding=options.encoding) as f:
            f.write(content)

    @staticmethod
    def save_markdown_to_string(workbook, options: Optional[MarkdownSaveOptions] = None) -> str:
        """
        Saves a workbook to a Markdown string.

        Args:
            workbook: The Workbook object to export.
            options (MarkdownSaveOptions, optional): Export options. Uses defaults if None.

        Returns:
            str: The Markdown content as a string.

        Examples:
            >>> wb = Workbook('data.xlsx')
            >>> md_content = MarkdownHandler.save_markdown_to_string(wb)
        """
        if options is None:
            options = MarkdownSaveOptions()

        output = io.StringIO()

        # Determine which worksheets to export
        if options.worksheet_index == -1:
            # Export all worksheets
            worksheets = [(i, ws) for i, ws in enumerate(workbook.worksheets)]
        else:
            if options.worksheet_index >= len(workbook.worksheets):
                raise IndexError(f"Worksheet index {options.worksheet_index} out of range")
            worksheets = [(options.worksheet_index, workbook.worksheets[options.worksheet_index])]

        for idx, (ws_idx, worksheet) in enumerate(worksheets):
            if idx > 0:
                output.write('\n\n')

            # Write worksheet name as header if enabled
            if options.include_worksheet_name:
                header_prefix = '#' * options.header_level
                output.write(f"{header_prefix} {worksheet.name}\n\n")

            # Get worksheet data
            rows_data = MarkdownHandler._get_worksheet_data(worksheet)

            if not rows_data:
                output.write('*No data*\n')
                continue

            # Process title rows and create tables
            if options.detect_title_rows or options.auto_detect_header:
                content = MarkdownHandler._create_markdown_with_titles(rows_data, options)
            else:
                content = MarkdownHandler._create_markdown_table(rows_data, options)
            output.write(content)

        return output.getvalue()

    @staticmethod
    def _get_worksheet_data(worksheet) -> List[List[Any]]:
        """
        Extracts all cell data from a worksheet as a 2D list.

        Args:
            worksheet: The Worksheet object to extract data from.

        Returns:
            list: 2D list of cell values, organized by rows.
        """
        from .cells import Cells

        cells_dict = worksheet.cells._cells

        if not cells_dict:
            return []

        # Find the dimensions by parsing cell references
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

        # Build 2D array
        rows_data = []
        for row_idx in range(1, max_row + 1):
            row_data = []
            for col_idx in range(1, max_col + 1):
                ref = Cells.coordinate_to_string(row_idx, col_idx)
                cell = cells_dict.get(ref)
                value = cell.value if cell else None
                row_data.append(value)
            rows_data.append(row_data)

        return rows_data

    @staticmethod
    def _is_title_row(row: List[Any]) -> bool:
        """
        Check if a row is a title row (only first cell has data).

        Args:
            row: List of cell values.

        Returns:
            bool: True if row has only first cell with data.
        """
        if not row:
            return False
        # Check if only the first cell has a non-empty value
        first_has_value = row[0] is not None and str(row[0]).strip() != ''
        rest_empty = all(val is None or str(val).strip() == '' for val in row[1:])
        return first_has_value and rest_empty

    @staticmethod
    def _is_empty_row(row: List[Any]) -> bool:
        """
        Check if a row is completely empty.

        Args:
            row: List of cell values.

        Returns:
            bool: True if all cells are empty.
        """
        return all(val is None or str(val).strip() == '' for val in row)

    @staticmethod
    def _create_markdown_with_titles(rows_data: List[List[Any]], options: MarkdownSaveOptions) -> str:
        """
        Creates Markdown with title rows converted to headings.

        Args:
            rows_data: 2D list of cell values.
            options: Export options.

        Returns:
            str: Markdown content with headings and tables.
        """
        if not rows_data:
            return ''

        output = io.StringIO()
        current_table_rows = []

        def flush_table():
            """Output accumulated table rows."""
            nonlocal current_table_rows
            if current_table_rows:
                # Skip leading empty rows
                while current_table_rows and MarkdownHandler._is_empty_row(current_table_rows[0]):
                    current_table_rows.pop(0)
                if current_table_rows:
                    table_md = MarkdownHandler._create_markdown_table(current_table_rows, options)
                    output.write(table_md)
                current_table_rows = []

        for row in rows_data:
            is_empty = MarkdownHandler._is_empty_row(row)
            is_title = MarkdownHandler._is_title_row(row)

            if is_empty:
                if options.skip_empty_rows:
                    continue
                else:
                    current_table_rows.append(row)
            elif is_title and options.detect_title_rows:
                # Flush any pending table
                flush_table()
                # Output title as heading (one level below worksheet name)
                title_text = MarkdownHandler._format_value(row[0], options)
                heading_level = options.header_level + 1
                heading_prefix = '#' * heading_level
                output.write(f"\n{heading_prefix} {title_text}\n\n")
            else:
                current_table_rows.append(row)

        # Flush remaining table
        flush_table()

        # Clean up multiple consecutive blank lines
        result = output.getvalue()
        while '\n\n\n' in result:
            result = result.replace('\n\n\n', '\n\n')
        return result.lstrip('\n')  # Remove leading newlines

    @staticmethod
    def _create_markdown_table(rows_data: List[List[Any]], options: MarkdownSaveOptions) -> str:
        """
        Creates a Markdown table from 2D data.

        Args:
            rows_data: 2D list of cell values.
            options: Export options.

        Returns:
            str: Markdown table string.
        """
        if not rows_data:
            return ''

        # Format all values
        formatted_rows = []
        for row_idx, row in enumerate(rows_data):
            formatted_row = [
                MarkdownHandler._format_value(val, options)
                for val in row
            ]

            # Skip empty rows if option is enabled (but always keep header row)
            if options.skip_empty_rows and row_idx > 0:
                if all(cell == options.empty_cell_placeholder or cell == '' for cell in formatted_row):
                    continue

            formatted_rows.append(formatted_row)

        # Calculate column widths
        num_cols = max(len(row) for row in formatted_rows) if formatted_rows else 0

        # Add row numbers column if enabled
        if options.include_row_numbers:
            start_row = 0 if options.first_row_as_header else 1
            for i, row in enumerate(formatted_rows):
                if i == 0 and options.first_row_as_header:
                    row.insert(0, '#')
                else:
                    row.insert(0, str(i + start_row))
            num_cols += 1

        # Ensure all rows have the same number of columns
        for row in formatted_rows:
            while len(row) < num_cols:
                row.append(options.empty_cell_placeholder)

        # Calculate column widths for alignment
        col_widths = [3] * num_cols  # Minimum width of 3 for separator
        for row in formatted_rows:
            for i, val in enumerate(row):
                width = len(val)
                if options.max_column_width > 0:
                    width = min(width, options.max_column_width)
                col_widths[i] = max(col_widths[i], width)

        # Build the table
        lines = []

        # Determine if using compact format
        use_compact = options.compact_format

        # Header row
        if options.first_row_as_header and formatted_rows:
            header_row = formatted_rows[0]
            if use_compact:
                header_cells = header_row
            else:
                header_cells = [
                    MarkdownHandler._pad_cell(val, col_widths[i], options, i)
                    for i, val in enumerate(header_row)
                ]
            lines.append('| ' + ' | '.join(header_cells) + ' |')

            # Separator row
            separator_cells = [
                MarkdownHandler._create_separator(col_widths[i], options, i)
                for i in range(num_cols)
            ]
            lines.append('| ' + ' | '.join(separator_cells) + ' |')

            # Data rows
            data_rows = formatted_rows[1:]
        else:
            # Create generic header
            header_cells = [f'Column {i+1}' for i in range(num_cols)]
            if not use_compact:
                header_cells = [
                    MarkdownHandler._pad_cell(val, col_widths[i], options, i)
                    for i, val in enumerate(header_cells)
                ]
            lines.append('| ' + ' | '.join(header_cells) + ' |')

            # Separator row
            separator_cells = [
                MarkdownHandler._create_separator(col_widths[i], options, i)
                for i in range(num_cols)
            ]
            lines.append('| ' + ' | '.join(separator_cells) + ' |')

            data_rows = formatted_rows

        # Data rows
        for row in data_rows:
            if use_compact:
                data_cells = row
            else:
                data_cells = [
                    MarkdownHandler._pad_cell(val, col_widths[i], options, i)
                    for i, val in enumerate(row)
                ]
            lines.append('| ' + ' | '.join(data_cells) + ' |')

        return '\n'.join(lines) + '\n'

    @staticmethod
    def _format_value(value: Any, options: MarkdownSaveOptions) -> str:
        """
        Formats a cell value for Markdown output.

        Args:
            value: The cell value to format.
            options: Export options.

        Returns:
            str: The formatted string value.
        """
        if value is None:
            return options.empty_cell_placeholder

        # Handle datetime types
        if isinstance(value, datetime):
            result = value.strftime(options.datetime_format)
        elif isinstance(value, date):
            result = value.strftime(options.date_format)
        elif isinstance(value, time):
            result = value.strftime(options.time_format)
        elif isinstance(value, bool):
            result = 'Yes' if value else 'No'
        elif isinstance(value, float):
            # Check if it's an integer stored as float
            if value.is_integer():
                result = str(int(value))
            elif options.float_precision >= 0:
                # Use fixed precision
                result = f"{value:.{options.float_precision}f}"
            else:
                # Auto precision: round to reasonable precision and strip trailing zeros
                # Use 10 significant digits to avoid floating point artifacts
                result = f"{value:.10g}"
        else:
            result = str(value)

        # Replace newlines to prevent breaking table structure
        if options.newline_replacement is not None:
            result = result.replace('\r\n', options.newline_replacement)
            result = result.replace('\n', options.newline_replacement)
            result = result.replace('\r', options.newline_replacement)

        # Trim whitespace if enabled
        if options.trim_whitespace:
            result = result.strip()

        # Escape pipe characters if enabled
        if options.escape_pipes:
            result = result.replace('|', '\\|')

        # Truncate if max width is set
        if options.max_column_width > 0 and len(result) > options.max_column_width:
            result = result[:options.max_column_width - 3] + '...'

        return result

    @staticmethod
    def _get_alignment(col_index: int, options: MarkdownSaveOptions) -> str:
        """
        Gets the alignment for a specific column.

        Args:
            col_index: Column index (0-based).
            options: Export options.

        Returns:
            str: Alignment string ('left', 'center', or 'right').
        """
        return options.column_alignments.get(col_index, options.default_alignment)

    @staticmethod
    def _create_separator(width: int, options: MarkdownSaveOptions, col_index: int) -> str:
        """
        Creates a separator cell for the Markdown table.

        Args:
            width: Column width.
            options: Export options.
            col_index: Column index (0-based).

        Returns:
            str: Separator string with alignment markers.
        """
        # Simple separators: just '---' without alignment markers
        if options.simple_separators:
            return '---'

        alignment = MarkdownHandler._get_alignment(col_index, options)

        # In compact format, use minimal width for separators but still show alignment
        if options.compact_format:
            if alignment == 'center':
                return ':---:'
            elif alignment == 'right':
                return '---:'
            else:  # left (default)
                return '---'

        # Non-compact format: use full width for separators
        if alignment == 'center':
            return ':' + '-' * width + ':'
        elif alignment == 'right':
            return '-' * (width + 1) + ':'
        else:  # left (default)
            return ':' + '-' * (width + 1)

    @staticmethod
    def _pad_cell(value: str, width: int, options: MarkdownSaveOptions, col_index: int) -> str:
        """
        Pads a cell value to the specified width with appropriate alignment.

        Args:
            value: Cell value string.
            width: Target width.
            options: Export options.
            col_index: Column index (0-based).

        Returns:
            str: Padded cell value.
        """
        alignment = MarkdownHandler._get_alignment(col_index, options)

        if len(value) >= width:
            return value

        if alignment == 'center':
            return value.center(width)
        elif alignment == 'right':
            return value.rjust(width)
        else:  # left (default)
            return value.ljust(width)


def save_workbook_as_markdown(workbook, file_path: str, options: Optional[MarkdownSaveOptions] = None) -> None:
    """
    Convenience function to save a Workbook to a Markdown file.

    Args:
        workbook: The Workbook object to export.
        file_path (str): Path where the Markdown file should be saved.
        options (MarkdownSaveOptions, optional): Export options. Uses defaults if None.

    Examples:
        >>> wb = Workbook('data.xlsx')
        >>> save_workbook_as_markdown(wb, 'output.md')
    """
    MarkdownHandler.save_markdown(workbook, file_path, options)
