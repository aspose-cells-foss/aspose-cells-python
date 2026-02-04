"""
Aspose.Cells for Python - AutoFilter Module

This module provides the AutoFilter class which represents auto filters in Excel worksheets.
The AutoFilter class provides methods for applying and managing filters on data ranges.

Compatible with Aspose.Cells for .NET API structure.
"""


class FilterColumn:
    """
    Represents a filter column in an auto filter.
    
    A FilterColumn represents filtering settings for a specific column in the filtered range.
    """
    
    def __init__(self, col_id):
        """
        Initializes a new instance of the FilterColumn class.
        
        Args:
            col_id (int): Zero-based column index in the filter range.
        """
        self._col_id = col_id
        self._filters = []  # List of filter values
        self._custom_filters = []  # List of custom filter criteria
        self._color_filter = None  # Color filter settings
        self._dynamic_filter = None  # Dynamic filter settings
        self._top10_filter = None  # Top 10 filter settings
        self._filter_button = True  # Whether filter button is visible
    
    @property
    def col_id(self):
        """
        Gets the zero-based column index.
        
        Returns:
            int: Zero-based column index.
        """
        return self._col_id
    
    @property
    def filters(self):
        """
        Gets the list of filter values.
        
        Returns:
            list: List of filter values.
        """
        return self._filters
    
    @property
    def custom_filters(self):
        """
        Gets the list of custom filter criteria.
        
        Returns:
            list: List of custom filter criteria tuples (operator, value).
        """
        return self._custom_filters
    
    @property
    def color_filter(self):
        """
        Gets or sets the color filter settings.
        
        Returns:
            dict or None: Color filter settings with keys 'color' and 'cell_color' (bool).
        """
        return self._color_filter
    
    @color_filter.setter
    def color_filter(self, value):
        """
        Sets the color filter settings.
        
        Args:
            value (dict): Color filter settings with keys 'color' and 'cell_color' (bool).
        """
        self._color_filter = value
    
    @property
    def dynamic_filter(self):
        """
        Gets or sets the dynamic filter settings.
        
        Returns:
            dict or None: Dynamic filter settings with keys 'type' and 'value'.
        """
        return self._dynamic_filter
    
    @dynamic_filter.setter
    def dynamic_filter(self, value):
        """
        Sets the dynamic filter settings.
        
        Args:
            value (dict): Dynamic filter settings with keys 'type' and 'value'.
        """
        self._dynamic_filter = value
    
    @property
    def top10_filter(self):
        """
        Gets or sets the top 10 filter settings.
        
        Returns:
            dict or None: Top 10 filter settings with keys 'top', 'percent', and 'val'.
        """
        return self._top10_filter
    
    @top10_filter.setter
    def top10_filter(self, value):
        """
        Sets the top 10 filter settings.
        
        Args:
            value (dict): Top 10 filter settings with keys 'top', 'percent', and 'val'.
        """
        self._top10_filter = value
    
    @property
    def filter_button(self):
        """
        Gets or sets whether the filter button is visible.
        
        Returns:
            bool: True if filter button is visible, False otherwise.
        """
        return self._filter_button
    
    @filter_button.setter
    def filter_button(self, value):
        """
        Sets whether the filter button is visible.
        
        Args:
            value (bool): True to show filter button, False to hide.
        """
        self._filter_button = value
    
    def add_filter(self, value):
        """
        Adds a filter value to this column.
        
        Args:
            value: The value to filter by.
            
        Examples:
            >>> filter_col.add_filter("Apple")
            >>> filter_col.add_filter(100)
        """
        self._filters.append(value)
    
    def add_custom_filter(self, operator, value):
        """
        Adds a custom filter criterion to this column.
        
        Args:
            operator (str): The operator ('equal', 'notEqual', 'greaterThan', 'lessThan', 
                         'greaterThanOrEqual', 'lessThanOrEqual', 'contains', 'notContains',
                         'beginsWith', 'endsWith').
            value: The value to compare against.
            
        Examples:
            >>> filter_col.add_custom_filter('greaterThan', 50)
            >>> filter_col.add_custom_filter('contains', 'test')
        """
        self._custom_filters.append((operator, value))
    
    def clear_filters(self):
        """
        Clears all filters from this column.
        
        Examples:
            >>> filter_col.clear_filters()
        """
        self._filters = []
        self._custom_filters = []
        self._color_filter = None
        self._dynamic_filter = None
        self._top10_filter = None


class AutoFilter:
    """
    Represents auto filters in a worksheet.
    
    The AutoFilter class provides methods and properties for applying and managing
    filters on data ranges in a worksheet.
    
    Examples:
        >>> from aspose_cells import Workbook
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> ws.cells['A1'].value = "Name"
        >>> ws.cells['B1'].value = "Age"
        >>> ws.auto_filter.range = "A1:B10"
        >>> ws.auto_filter.filter(0, ["Alice", "Bob"])
    """
    
    def __init__(self):
        """
        Initializes a new instance of the AutoFilter class.
        """
        self._range = None  # Filter range in A1 notation (e.g., "A1:D10")
        self._filter_columns = {}  # Dictionary mapping col_id to FilterColumn objects
        self._sort_state = None  # Sort state settings
    
    @property
    def range(self):
        """
        Gets or sets the filter range.
        
        Returns:
            str or None: Filter range in A1 notation (e.g., "A1:D10").
            
        Examples:
            >>> ws.auto_filter.range = "A1:D10"
            >>> print(ws.auto_filter.range)
        """
        return self._range
    
    @range.setter
    def range(self, value):
        """
        Sets the filter range.
        
        Args:
            value (str): Filter range in A1 notation (e.g., "A1:D10").
            
        Examples:
            >>> ws.auto_filter.range = "A1:D10"
        """
        self._range = value
    
    @property
    def filter_columns(self):
        """
        Gets the collection of filter columns.
        
        Returns:
            dict: Dictionary mapping col_id to FilterColumn objects.
        """
        return self._filter_columns
    
    @property
    def sort_state(self):
        """
        Gets or sets the sort state settings.
        
        Returns:
            dict or None: Sort state settings with keys 'column_offset', 'sort_order', etc.
        """
        return self._sort_state
    
    @sort_state.setter
    def sort_state(self, value):
        """
        Sets the sort state settings.
        
        Args:
            value (dict): Sort state settings with keys 'column_offset', 'sort_order', etc.
        """
        self._sort_state = value
    
    def set_range(self, start_row, start_col, end_row, end_col):
        """
        Sets the filter range using row and column indices.
        
        Args:
            start_row (int): 1-based starting row number.
            start_col (int or str): 1-based starting column number or letter (e.g., 1 or 'A').
            end_row (int): 1-based ending row number.
            end_col (int or str): 1-based ending column number or letter (e.g., 4 or 'D').
            
        Examples:
            >>> ws.auto_filter.set_range(1, 1, 10, 4)  # A1:D10
            >>> ws.auto_filter.set_range(1, 'A', 10, 'D')  # A1:D10
        """
        from .cells import Cells
        
        # Convert column to letter if necessary
        if isinstance(start_col, int):
            start_col_letter = Cells.column_letter_from_index(start_col)
        else:
            start_col_letter = start_col.upper()
        
        if isinstance(end_col, int):
            end_col_letter = Cells.column_letter_from_index(end_col)
        else:
            end_col_letter = end_col.upper()
        
        self._range = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"
    
    def filter(self, col_index, values):
        """
        Applies a filter to a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            values (list): List of values to filter by.
            
        Examples:
            >>> ws.auto_filter.filter(0, ["Apple", "Banana"])  # Filter first column
            >>> ws.auto_filter.filter(1, [10, 20, 30])  # Filter second column
        """
        if col_index not in self._filter_columns:
            self._filter_columns[col_index] = FilterColumn(col_index)
        
        filter_col = self._filter_columns[col_index]
        filter_col._filters = list(values)
    
    def add_filter(self, col_index, value):
        """
        Adds a filter value to a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            value: The value to filter by.
            
        Examples:
            >>> ws.auto_filter.add_filter(0, "Apple")
            >>> ws.auto_filter.add_filter(1, 100)
        """
        if col_index not in self._filter_columns:
            self._filter_columns[col_index] = FilterColumn(col_index)
        
        self._filter_columns[col_index].add_filter(value)
    
    def custom_filter(self, col_index, operator, value):
        """
        Applies a custom filter to a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            operator (str): The operator ('equal', 'notEqual', 'greaterThan', 'lessThan',
                         'greaterThanOrEqual', 'lessThanOrEqual', 'contains', 'notContains',
                         'beginsWith', 'endsWith').
            value: The value to compare against.
            
        Examples:
            >>> ws.auto_filter.custom_filter(0, 'greaterThan', 50)
            >>> ws.auto_filter.custom_filter(1, 'contains', 'test')
        """
        if col_index not in self._filter_columns:
            self._filter_columns[col_index] = FilterColumn(col_index)
        
        self._filter_columns[col_index].add_custom_filter(operator, value)
    
    def filter_by_color(self, col_index, color, cell_color=True):
        """
        Applies a color filter to a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            color (str): RGB hex color string in AARRGGBB format.
            cell_color (bool): True to filter by cell color, False to filter by font color.
            
        Examples:
            >>> ws.auto_filter.filter_by_color(0, 'FFFF0000')  # Filter by red cell color
            >>> ws.auto_filter.filter_by_color(1, 'FF0000FF', False)  # Filter by blue font color
        """
        if col_index not in self._filter_columns:
            self._filter_columns[col_index] = FilterColumn(col_index)
        
        self._filter_columns[col_index].color_filter = {
            'color': color,
            'cell_color': cell_color
        }
    
    def filter_top10(self, col_index, top=True, percent=False, val=10):
        """
        Applies a top 10 filter to a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            top (bool): True for top items, False for bottom items.
            percent (bool): True for percentage, False for count.
            val (int): The value (count or percentage).
            
        Examples:
            >>> ws.auto_filter.filter_top10(0)  # Top 10 items
            >>> ws.auto_filter.filter_top10(1, top=False)  # Bottom 10 items
            >>> ws.auto_filter.filter_top10(2, percent=True, val=20)  # Top 20%
        """
        if col_index not in self._filter_columns:
            self._filter_columns[col_index] = FilterColumn(col_index)
        
        self._filter_columns[col_index].top10_filter = {
            'top': top,
            'percent': percent,
            'val': val
        }
    
    def filter_dynamic(self, col_index, filter_type, value=None):
        """
        Applies a dynamic filter to a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            filter_type (str): The dynamic filter type ('aboveAverage', 'belowAverage',
                             'lastMonth', 'lastQuarter', 'lastWeek', 'lastYear',
                             'nextMonth', 'nextQuarter', 'nextWeek', 'nextYear',
                             'thisMonth', 'thisQuarter', 'thisWeek', 'thisYear',
                             'today', 'tomorrow', 'yesterday', 'yearToDate').
            value: Optional value for the filter.
            
        Examples:
            >>> ws.auto_filter.filter_dynamic(0, 'aboveAverage')
            >>> ws.auto_filter.filter_dynamic(1, 'lastMonth')
        """
        if col_index not in self._filter_columns:
            self._filter_columns[col_index] = FilterColumn(col_index)
        
        self._filter_columns[col_index].dynamic_filter = {
            'type': filter_type,
            'value': value
        }
    
    def clear_column_filter(self, col_index):
        """
        Clears the filter for a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            
        Examples:
            >>> ws.auto_filter.clear_column_filter(0)
        """
        if col_index in self._filter_columns:
            self._filter_columns[col_index].clear_filters()
    
    def clear_all_filters(self):
        """
        Clears all filters.
        
        Examples:
            >>> ws.auto_filter.clear_all_filters()
        """
        self._filter_columns = {}
    
    def remove(self):
        """
        Removes the auto filter from the worksheet.
        
        Examples:
            >>> ws.auto_filter.remove()
        """
        self._range = None
        self._filter_columns = {}
        self._sort_state = None
    
    def show_filter_button(self, col_index, show=True):
        """
        Shows or hides the filter button for a specific column.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            show (bool): True to show, False to hide.
            
        Examples:
            >>> ws.auto_filter.show_filter_button(0, False)  # Hide filter button
        """
        if col_index not in self._filter_columns:
            self._filter_columns[col_index] = FilterColumn(col_index)
        
        self._filter_columns[col_index].filter_button = show
    
    def sort(self, col_index, ascending=True):
        """
        Sets the sort order for a specific column.

        Args:
            col_index (int): Zero-based column index within the filter range.
            ascending (bool): True for ascending order, False for descending.

        Examples:
            >>> ws.auto_filter.sort(0, True)  # Sort first column ascending
            >>> ws.auto_filter.sort(1, False)  # Sort second column descending
        """
        self._sort_state = {
            'column_index': col_index,
            'descending': not ascending
        }
    
    def get_filter_column(self, col_index):
        """
        Gets the filter column for a specific column index.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            
        Returns:
            FilterColumn or None: The FilterColumn object, or None if not found.
            
        Examples:
            >>> filter_col = ws.auto_filter.get_filter_column(0)
            >>> if filter_col:
            ...     print(filter_col.filters)
        """
        return self._filter_columns.get(col_index)
    
    def has_filter(self, col_index):
        """
        Checks if a specific column has filters applied.
        
        Args:
            col_index (int): Zero-based column index within the filter range.
            
        Returns:
            bool: True if the column has filters, False otherwise.
            
        Examples:
            >>> if ws.auto_filter.has_filter(0):
            ...     print("Column 0 has filters")
        """
        if col_index not in self._filter_columns:
            return False
        
        filter_col = self._filter_columns[col_index]
        return (len(filter_col.filters) > 0 or 
                len(filter_col.custom_filters) > 0 or
                filter_col.color_filter is not None or
                filter_col.dynamic_filter is not None or
                filter_col.top10_filter is not None)
