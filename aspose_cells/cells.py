"""
Aspose.Cells for Python - Cells Module

This module provides the Cells class which represents a collection of cells in a worksheet.
The Cells class provides methods for accessing, modifying, and iterating over cells.

Compatible with Aspose.Cells for .NET API structure.
"""

from .cell import Cell


class Cells:
    """
    Represents a collection of cells in a worksheet.
    
    The Cells class provides methods and properties for working with collections of cells,
    including cell access by reference or coordinates, iteration, and coordinate conversion.
    
    Examples:
        >>> from aspose_cells import Workbook
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> cells = ws.cells
        >>> cells['A1'].value = "Hello"
        >>> cell = cells.cell(row=1, column=1)
    """
    
    def __init__(self, worksheet=None):
        """
        Initializes a new instance of the Cells class.
        
        Examples:
            >>> cells = Cells()
        """
        self._cells = {}
        self._worksheet = worksheet

    def _require_worksheet(self):
        if self._worksheet is None:
            raise ValueError("Cells is not attached to a Worksheet")
    
    # Cell access methods
    
    def __getitem__(self, key):
        """
        Gets a cell by its reference (e.g., 'A1').
        
        Args:
            key (str): Cell reference in A1 notation (e.g., 'A1', 'B3').
            
        Returns:
            Cell: The Cell object at the specified reference. Creates a new Cell if it doesn't exist.
            
        Examples:
            >>> cell = cells['A1']
            >>> cell.value = "Hello"
        """
        if key not in self._cells:
            self._cells[key] = Cell()
        return self._cells[key]
    
    def __setitem__(self, key, value):
        """
        Sets a cell value by its reference (e.g., 'A1').
        
        Args:
            key (str): Cell reference in A1 notation (e.g., 'A1', 'B3').
            value: The value to set in the cell. Can be a Cell object or any value type.
            
        Examples:
            >>> cells['A1'] = "Hello"
            >>> cells['B1'] = 42
            >>> cells['C1'] = Cell("Custom cell")
        """
        if key not in self._cells:
            self._cells[key] = Cell()
        if isinstance(value, Cell):
            self._cells[key] = value
        else:
            self._cells[key].value = value
    
    def cell(self, row=None, column=None):
        """
        Accesses a cell by row and column (1-based).
        
        Args:
            row (int): 1-based row number.
            column (int or str): 1-based column number or column letter (e.g., 1 or 'A').
            
        Returns:
            Cell: The Cell object at the specified row and column.
            
        Raises:
            ValueError: If row or column is not specified.
            
        Examples:
            >>> cell = cells.cell(row=1, column=1)  # Same as cells['A1']
            >>> cell = cells.cell(row=3, column='B')  # Same as cells['B3']
            >>> cell = cells.cell(row=5, column=10)  # Cell J5
        """
        if row is None or column is None:
            raise ValueError("Both row and column must be specified")
        
        # Convert column letter to number if needed
        if isinstance(column, str):
            column = self.column_index_from_string(column)
        
        ref = f"{self.column_letter_from_index(column)}{row}"
        return self[ref]
    
    # Coordinate conversion methods (static methods)
    
    @staticmethod
    def column_index_from_string(column):
        """
        Converts a column letter to a 1-based index.
        
        Args:
            column (str): Column letter(s) (e.g., 'A', 'Z', 'AA').
            
        Returns:
            int: 1-based column index (e.g., 'A' -> 1, 'Z' -> 26, 'AA' -> 27).
            
        Raises:
            ValueError: If column string is empty or contains invalid characters.
            
        Examples:
            >>> Cells.column_index_from_string('A')
            1
            >>> Cells.column_index_from_string('Z')
            26
            >>> Cells.column_index_from_string('AA')
            27
            >>> Cells.column_index_from_string('AB')
            28
        """
        if not column:
            raise ValueError("Column string cannot be empty")
        
        column = column.upper()
        result = 0
        for char in column:
            if not char.isalpha():
                raise ValueError(f"Invalid column character: {char}")
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    @staticmethod
    def column_letter_from_index(column_index):
        """
        Converts a 1-based column index to a letter.
        
        Args:
            column_index (int): 1-based column index.
            
        Returns:
            str: Column letter(s) (e.g., 1 -> 'A', 26 -> 'Z', 27 -> 'AA').
            
        Raises:
            ValueError: If column_index is less than 1.
            
        Examples:
            >>> Cells.column_letter_from_index(1)
            'A'
            >>> Cells.column_letter_from_index(26)
            'Z'
            >>> Cells.column_letter_from_index(27)
            'AA'
            >>> Cells.column_letter_from_index(28)
            'AB'
        """
        if column_index < 1:
            raise ValueError("Column index must be >= 1")
        
        result = ""
        while column_index > 0:
            column_index -= 1
            result = chr(ord('A') + (column_index % 26)) + result
            column_index = column_index // 26
        return result
    
    @staticmethod
    def coordinate_from_string(coord):
        """
        Converts an A1 coordinate string to (row, column) tuple (1-based).
        
        Args:
            coord (str): Cell reference in A1 notation (e.g., 'A1', 'B3', 'Z100').
            
        Returns:
            tuple: (row, column) tuple with 1-based indices.
            
        Raises:
            ValueError: If coordinate is empty or has invalid format.
            
        Examples:
            >>> Cells.coordinate_from_string('A1')
            (1, 1)
            >>> Cells.coordinate_from_string('B3')
            (3, 2)
            >>> Cells.coordinate_from_string('AA10')
            (10, 27)
        """
        if not coord:
            raise ValueError("Coordinate cannot be empty")
        
        # Split into column letters and row number
        col_str = ""
        row_str = ""
        for char in coord:
            if char.isalpha():
                col_str += char
            elif char.isdigit():
                row_str += char
            else:
                raise ValueError(f"Invalid character in coordinate: {char}")
        
        if not col_str or not row_str:
            raise ValueError(f"Invalid coordinate format: {coord}")
        
        column = Cells.column_index_from_string(col_str)
        row = int(row_str)
        return (row, column)
    
    @staticmethod
    def coordinate_to_string(row, column):
        """
        Converts row and column (1-based) to an A1 coordinate string.
        
        Args:
            row (int): 1-based row number.
            column (int): 1-based column number.
            
        Returns:
            str: Cell reference in A1 notation.
            
        Raises:
            ValueError: If row or column is less than 1.
            
        Examples:
            >>> Cells.coordinate_to_string(1, 1)
            'A1'
            >>> Cells.coordinate_to_string(3, 2)
            'B3'
            >>> Cells.coordinate_to_string(10, 27)
            'AA10'
        """
        if row < 1 or column < 1:
            raise ValueError("Row and column must be >= 1")
        
        col_letter = Cells.column_letter_from_index(column)
        return f"{col_letter}{row}"
    
    # Iteration methods
    
    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False):
        """
        Iterates over rows in the worksheet.
        
        Args:
            min_row (int, optional): Minimum row number (1-based). Defaults to 1.
            max_row (int, optional): Maximum row number (1-based). Defaults to max row with data.
            min_col (int, optional): Minimum column number (1-based). Defaults to 1.
            max_col (int, optional): Maximum column number (1-based). Defaults to max column with data.
            values_only (bool): If True, yields cell values instead of Cell objects. Defaults to False.
            
        Yields:
            tuple: If values_only=False, yields tuples containing Cell objects for each row.
                    If values_only=True, yields tuples containing cell values for each row.
                    
        Examples:
            >>> for row in cells.iter_rows(min_row=1, max_row=5):
            ...     print(row[0].value)
            
            >>> for row in cells.iter_rows(min_row=1, max_row=3, min_col=1, max_col=3, values_only=True):
            ...     print(row)
        """
        # Determine bounds
        if not self._cells:
            return
        
        # Get all row and column numbers from existing cells
        rows = set()
        cols = set()
        for ref in self._cells:
            row, col = self.coordinate_from_string(ref)
            rows.add(row)
            cols.add(col)
        
        min_row = min_row if min_row is not None else min(rows) if rows else 1
        max_row = max_row if max_row is not None else max(rows) if rows else 1
        min_col = min_col if min_col is not None else min(cols) if cols else 1
        max_col = max_col if max_col is not None else max(cols) if cols else 1
        
        for row in range(min_row, max_row + 1):
            row_cells = []
            for col in range(min_col, max_col + 1):
                ref = self.coordinate_to_string(row, col)
                cell = self[ref]
                if values_only:
                    row_cells.append(cell.value)
                else:
                    row_cells.append(cell)
            yield tuple(row_cells)
    
    def iter_cols(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False):
        """
        Iterates over columns in the worksheet.
        
        Args:
            min_row (int, optional): Minimum row number (1-based). Defaults to 1.
            max_row (int, optional): Maximum row number (1-based). Defaults to max row with data.
            min_col (int, optional): Minimum column number (1-based). Defaults to 1.
            max_col (int, optional): Maximum column number (1-based). Defaults to max column with data.
            values_only (bool): If True, yields cell values instead of Cell objects. Defaults to False.
            
        Yields:
            tuple: If values_only=False, yields tuples containing Cell objects for each column.
                    If values_only=True, yields tuples containing cell values for each column.
                    
        Examples:
            >>> for col in cells.iter_cols(min_col=1, max_col=3):
            ...     print(col[0].value)
            
            >>> for col in cells.iter_cols(min_row=1, max_row=5, min_col=1, max_col=2, values_only=True):
            ...     print(col)
        """
        # Determine bounds
        if not self._cells:
            return
        
        # Get all row and column numbers from existing cells
        rows = set()
        cols = set()
        for ref in self._cells:
            row, col = self.coordinate_from_string(ref)
            rows.add(row)
            cols.add(col)
        
        min_row = min_row if min_row is not None else min(rows) if rows else 1
        max_row = max_row if max_row is not None else max(rows) if rows else 1
        min_col = min_col if min_col is not None else min(cols) if cols else 1
        max_col = max_col if max_col is not None else max(cols) if cols else 1
        
        for col in range(min_col, max_col + 1):
            col_cells = []
            for row in range(min_row, max_row + 1):
                ref = self.coordinate_to_string(row, col)
                cell = self[ref]
                if values_only:
                    col_cells.append(cell.value)
                else:
                    col_cells.append(cell)
            yield tuple(col_cells)
    
    # Cell collection methods
    
    def count(self):
        """
        Gets the number of cells in the collection.
        
        Returns:
            int: The number of cells that have been accessed or modified.
            
        Examples:
            >>> num_cells = cells.count()
            >>> print(f"Total cells: {num_cells}")
        """
        return len(self._cells)
    
    def clear(self):
        """
        Clears all cells in the collection.
        
        Examples:
            >>> cells.clear()
        """
        self._cells.clear()
    
    def get_cell_by_name(self, cell_name):
        """
        Gets a cell by its name (reference).
        
        Args:
            cell_name (str): Cell reference in A1 notation (e.g., 'A1', 'B3').
            
        Returns:
            Cell: The Cell object at the specified reference.
            
        Examples:
            >>> cell = cells.get_cell_by_name('A1')
            >>> cell.value = "Hello"
        """
        return self[cell_name]
    
    def set_cell_by_name(self, cell_name, value):
        """
        Sets a cell value by its name (reference).
        
        Args:
            cell_name (str): Cell reference in A1 notation (e.g., 'A1', 'B3').
            value: The value to set in the cell.
            
        Examples:
            >>> cells.set_cell_by_name('A1', "Hello")
            >>> cells.set_cell_by_name('B1', 42)
        """
        self[cell_name] = value
    
    def get_cell(self, row, column):
        """
        Gets a cell by row and column (1-based).
        
        Args:
            row (int): 1-based row number.
            column (int): 1-based column number.
            
        Returns:
            Cell: The Cell object at the specified row and column.
            
        Examples:
            >>> cell = cells.get_cell(1, 1)  # Same as cells['A1']
            >>> cell = cells.get_cell(3, 2)  # Same as cells['B3']
        """
        return self.cell(row=row, column=column)
    
    def set_cell(self, row, column, value):
        """
        Sets a cell value by row and column (1-based).
        
        Args:
            row (int): 1-based row number.
            column (int): 1-based column number.
            value: The value to set in the cell.
            
        Examples:
            >>> cells.set_cell(1, 1, "Hello")  # Same as cells['A1'] = "Hello"
            >>> cells.set_cell(3, 2, 42)  # Same as cells['B3'] = 42
        """
        ref = self.coordinate_to_string(row, column)
        self[ref] = value

    # Row/Column dimensions (Aspose.Cells compatible)

    def set_row_height(self, row, height):
        """
        Sets the height of the specified row in points.

        Args:
            row (int): 1-based row index.
            height (float): Row height in points (must be > 0).
        """
        self._require_worksheet()
        if row is None or row < 1:
            raise ValueError("row must be >= 1")
        if height is None or height <= 0:
            raise ValueError("height must be > 0")
        self._worksheet._row_heights[int(row)] = float(height)

    def get_row_height(self, row):
        """
        Gets the height of the specified row in points.

        Args:
            row (int): 1-based row index.

        Returns:
            float: Row height in points.
        """
        self._require_worksheet()
        if row is None or row < 1:
            raise ValueError("row must be >= 1")
        row = int(row)
        if row in self._worksheet._row_heights:
            return self._worksheet._row_heights[row]
        return float(self._worksheet.properties.format.default_row_height)

    def hide_row(self, row):
        """
        Hides the specified row.

        Args:
            row (int): 1-based row index.
        """
        self._require_worksheet()
        if row is None or row < 1:
            raise ValueError("row must be >= 1")
        self._worksheet._hidden_rows.add(int(row))

    def unhide_row(self, row):
        """
        Unhides the specified row.

        Args:
            row (int): 1-based row index.
        """
        self._require_worksheet()
        if row is None or row < 1:
            raise ValueError("row must be >= 1")
        self._worksheet._hidden_rows.discard(int(row))

    def is_row_hidden(self, row):
        """
        Checks if the specified row is hidden.

        Args:
            row (int): 1-based row index.

        Returns:
            bool: True if hidden, False otherwise.
        """
        self._require_worksheet()
        if row is None or row < 1:
            raise ValueError("row must be >= 1")
        return int(row) in self._worksheet._hidden_rows

    def set_column_width(self, column, width):
        """
        Sets the width of the specified column in character units.

        Args:
            column (int or str): 1-based column index or column letter.
            width (float): Column width in characters (must be > 0).
        """
        self._require_worksheet()
        if column is None:
            raise ValueError("column must be specified")
        if isinstance(column, str):
            column = self.column_index_from_string(column)
        if column < 1:
            raise ValueError("column must be >= 1")
        if width is None or width <= 0:
            raise ValueError("width must be > 0")
        self._worksheet._column_widths[int(column)] = float(width)

    def get_column_width(self, column):
        """
        Gets the width of the specified column in character units.

        Args:
            column (int or str): 1-based column index or column letter.

        Returns:
            float: Column width in characters.
        """
        self._require_worksheet()
        if column is None:
            raise ValueError("column must be specified")
        if isinstance(column, str):
            column = self.column_index_from_string(column)
        if column < 1:
            raise ValueError("column must be >= 1")
        column = int(column)
        if column in self._worksheet._column_widths:
            return self._worksheet._column_widths[column]
        fmt = self._worksheet.properties.format
        if fmt.default_col_width is not None:
            return float(fmt.default_col_width)
        return float(fmt.base_col_width)

    def hide_column(self, column):
        """
        Hides the specified column.

        Args:
            column (int or str): 1-based column index or column letter.
        """
        self._require_worksheet()
        if column is None:
            raise ValueError("column must be specified")
        if isinstance(column, str):
            column = self.column_index_from_string(column)
        if column < 1:
            raise ValueError("column must be >= 1")
        self._worksheet._hidden_columns.add(int(column))

    def unhide_column(self, column):
        """
        Unhides the specified column.

        Args:
            column (int or str): 1-based column index or column letter.
        """
        self._require_worksheet()
        if column is None:
            raise ValueError("column must be specified")
        if isinstance(column, str):
            column = self.column_index_from_string(column)
        if column < 1:
            raise ValueError("column must be >= 1")
        self._worksheet._hidden_columns.discard(int(column))

    def is_column_hidden(self, column):
        """
        Checks if the specified column is hidden.

        Args:
            column (int or str): 1-based column index or column letter.

        Returns:
            bool: True if hidden, False otherwise.
        """
        self._require_worksheet()
        if column is None:
            raise ValueError("column must be specified")
        if isinstance(column, str):
            column = self.column_index_from_string(column)
        if column < 1:
            raise ValueError("column must be >= 1")
        return int(column) in self._worksheet._hidden_columns

    # Aspose.Cells .NET-style aliases

    def SetRowHeight(self, row, height):
        return self.set_row_height(row, height)

    def GetRowHeight(self, row):
        return self.get_row_height(row)

    def SetColumnWidth(self, column, width):
        return self.set_column_width(column, width)

    def GetColumnWidth(self, column):
        return self.get_column_width(column)

    def SetRowHidden(self, row, is_hidden):
        if is_hidden:
            return self.hide_row(row)
        return self.unhide_row(row)

    def IsRowHidden(self, row):
        return self.is_row_hidden(row)

    def SetColumnHidden(self, column, is_hidden):
        if is_hidden:
            return self.hide_column(column)
        return self.unhide_column(column)

    def IsColumnHidden(self, column):
        return self.is_column_hidden(column)
    
    # Range methods
    
    def get_range(self, start_row, start_column, end_row, end_column):
        """
        Gets a range of cells as a list of lists.
        
        Args:
            start_row (int): 1-based start row number.
            start_column (int): 1-based start column number.
            end_row (int): 1-based end row number.
            end_column (int): 1-based end column number.
            
        Returns:
            list: List of lists containing Cell objects for the specified range.
            
        Examples:
            >>> range_cells = cells.get_range(1, 1, 3, 3)  # A1:C3 range
            >>> for row in range_cells:
            ...     for cell in row:
            ...         print(cell.value)
        """
        result = []
        for row in range(start_row, end_row + 1):
            row_cells = []
            for col in range(start_column, end_column + 1):
                ref = self.coordinate_to_string(row, col)
                row_cells.append(self[ref])
            result.append(row_cells)
        return result
    
    def set_range(self, start_row, start_column, end_row, end_column, values):
        """
        Sets values for a range of cells.
        
        Args:
            start_row (int): 1-based start row number.
            start_column (int): 1-based start column number.
            end_row (int): 1-based end row number.
            end_column (int): 1-based end column number.
            values: List of lists containing values to set. Must match the range dimensions.
            
        Examples:
            >>> values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
            >>> cells.set_range(1, 1, 3, 3, values)  # Sets A1:C3
        """
        for i, row in enumerate(range(start_row, end_row + 1)):
            for j, col in enumerate(range(start_column, end_column + 1)):
                if i < len(values) and j < len(values[i]):
                    ref = self.coordinate_to_string(row, col)
                    self[ref] = values[i][j]
    
    # Utility methods
    
    def has_cell(self, cell_name):
        """
        Checks if a cell exists in the collection.
        
        Args:
            cell_name (str): Cell reference in A1 notation (e.g., 'A1', 'B3').
            
        Returns:
            bool: True if the cell has been accessed or modified, False otherwise.
            
        Examples:
            >>> if cells.has_cell('A1'):
            ...     print("Cell A1 exists")
        """
        return cell_name in self._cells
    
    def delete_cell(self, cell_name):
        """
        Deletes a cell from the collection.
        
        Args:
            cell_name (str): Cell reference in A1 notation (e.g., 'A1', 'B3').
            
        Examples:
            >>> cells.delete_cell('A1')
        """
        if cell_name in self._cells:
            del self._cells[cell_name]
    
    def get_all_cells(self):
        """
        Gets all cells in the collection.
        
        Returns:
            dict: Dictionary mapping cell references to Cell objects.
            
        Examples:
            >>> all_cells = cells.get_all_cells()
            >>> for ref, cell in all_cells.items():
            ...     print(f"{ref}: {cell.value}")
        """
        return self._cells.copy()
    
    # String representation
    
    def __len__(self):
        """
        Gets the number of cells in the collection.
        
        Returns:
            int: The number of cells in the collection.
        """
        return len(self._cells)
    
    def __contains__(self, key):
        """
        Checks if a cell reference exists in the collection.
        
        Args:
            key (str): Cell reference in A1 notation.
            
        Returns:
            bool: True if the cell exists, False otherwise.
        """
        return key in self._cells
    
    def __iter__(self):
        """
        Iterates over all cells in the collection.
        
        Yields:
            tuple: (cell_reference, Cell) tuples for each cell.
        """
        for ref, cell in self._cells.items():
            yield (ref, cell)
    
    def __repr__(self):
        """
        Returns a string representation of the Cells collection.
        
        Returns:
            str: String representation showing the number of cells.
        """
        return f"Cells(count={len(self._cells)})"
