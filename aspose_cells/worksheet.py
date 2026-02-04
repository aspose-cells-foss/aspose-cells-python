"""
Aspose.Cells for Python - Worksheet Module

This module provides the Worksheet class which represents a single worksheet in an Excel workbook.
The Worksheet class provides methods to manage cells, ranges, and worksheet properties.

Compatible with Aspose.Cells for .NET API structure.
"""

from .cells import Cells
from .cell import Cell
from .auto_filter import AutoFilter
from .conditional_format import ConditionalFormatCollection
from .worksheet_properties import WorksheetProperties
from .hyperlink import Hyperlinks
from .data_validation import DataValidationCollection


class SheetProtectionDictWrapper:
    """
    Dictionary-like wrapper around SheetProtection for backward compatibility.

    This wrapper allows accessing SheetProtection properties using dictionary syntax
    while maintaining backward compatibility with existing code that uses ws.protection['key'].
    """

    def __init__(self, sheet_protection):
        """
        Initialize wrapper with a SheetProtection object.

        Args:
            sheet_protection: The SheetProtection object to wrap.
        """
        self._protection = sheet_protection

    def __getitem__(self, key):
        """Get protection setting by key."""
        if key == 'protected':
            return self._protection.sheet
        elif key == 'password':
            return self._protection.password
        elif key == 'sheet':
            return self._protection.sheet
        elif key == 'objects':
            return self._protection.objects
        elif key == 'scenarios':
            return self._protection.scenarios
        elif key == 'format_cells':
            return self._protection.format_cells
        elif key == 'format_columns':
            return self._protection.format_columns
        elif key == 'format_rows':
            return self._protection.format_rows
        elif key == 'insert_columns':
            return self._protection.insert_columns
        elif key == 'insert_rows':
            return self._protection.insert_rows
        elif key == 'insert_hyperlinks':
            return self._protection.insert_hyperlinks
        elif key == 'delete_columns':
            return self._protection.delete_columns
        elif key == 'delete_rows':
            return self._protection.delete_rows
        elif key == 'select_locked_cells':
            return self._protection.select_locked_cells
        elif key == 'select_unlocked_cells':
            return self._protection.select_unlocked_cells
        elif key == 'sort':
            return self._protection.sort
        elif key == 'auto_filter':
            return self._protection.auto_filter
        elif key == 'pivot_tables':
            return self._protection.pivot_tables
        else:
            raise KeyError(f"Unknown protection key: {key}")

    def __setitem__(self, key, value):
        """Set protection setting by key."""
        if key == 'protected':
            self._protection.sheet = value
        elif key == 'password':
            self._protection.password = value
        elif key == 'sheet':
            self._protection.sheet = value
        elif key == 'objects':
            self._protection.objects = value
        elif key == 'scenarios':
            self._protection.scenarios = value
        elif key == 'format_cells':
            self._protection.format_cells = value
        elif key == 'format_columns':
            self._protection.format_columns = value
        elif key == 'format_rows':
            self._protection.format_rows = value
        elif key == 'insert_columns':
            self._protection.insert_columns = value
        elif key == 'insert_rows':
            self._protection.insert_rows = value
        elif key == 'insert_hyperlinks':
            self._protection.insert_hyperlinks = value
        elif key == 'delete_columns':
            self._protection.delete_columns = value
        elif key == 'delete_rows':
            self._protection.delete_rows = value
        elif key == 'select_locked_cells':
            self._protection.select_locked_cells = value
        elif key == 'select_unlocked_cells':
            self._protection.select_unlocked_cells = value
        elif key == 'sort':
            self._protection.sort = value
        elif key == 'auto_filter':
            self._protection.auto_filter = value
        elif key == 'pivot_tables':
            self._protection.pivot_tables = value
        else:
            raise KeyError(f"Unknown protection key: {key}")

    def get(self, key, default=None):
        """Get protection setting with default value."""
        try:
            return self[key]
        except KeyError:
            return default


class Worksheet:
    """
    Represents a single worksheet in an Excel workbook.
    
    The Worksheet class provides access to cells, ranges, and worksheet properties
    such as name, visibility, protection, and page setup settings.
    
    Examples:
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> ws.name = "MySheet"
        >>> ws.cells["A1"].value = "Hello"
    """
    
    def __init__(self, name="Sheet1"):
        """
        Initializes a new instance of the Worksheet class.
        
        Args:
            name (str, optional): Name of the worksheet. Defaults to "Sheet1".
        """
        self._name = name
        self._cells = Cells(self)
        self._visible = True
        self._tab_color = None
        self._auto_filter = AutoFilter()  # Auto filter settings
        self._conditional_formats = ConditionalFormatCollection()  # Conditional formatting
        self._data_validations = DataValidationCollection()  # Data validation rules
        self._properties = WorksheetProperties()  # Worksheet properties
        self._hyperlinks = Hyperlinks(self)  # Hyperlinks collection
        self._row_heights = {}  # Row index -> height (points)
        self._column_widths = {}  # Column index -> width (characters)
        self._hidden_rows = set()  # Set of hidden row indices
        self._hidden_columns = set()  # Set of hidden column indices
        
        # Page setup settings
        self._page_setup = {
            'orientation': None,  # 'portrait' or 'landscape'
            'paper_size': None,  # Integer paper size
            'scale': None,  # Integer scale (10-400)
            'fit_to_width': None,  # Integer number of pages
            'fit_to_height': None,  # Integer number of pages
            'fit_to_page': False  # Boolean fit to page
        }
        
        # Page margins (in inches)
        self._page_margins = {
            'left': 0.75,
            'right': 0.75,
            'top': 1.0,
            'bottom': 1.0,
            'header': 0.5,
            'footer': 0.5
        }
    
    # Properties
    
    @property
    def name(self):
        """
        Gets or sets the name of the worksheet.
        
        Returns:
            str: The name of the worksheet.
        """
        return self._name
    
    @name.setter
    def name(self, value):
        """
        Sets the name of the worksheet.
        
        Args:
            value (str): The new name for the worksheet.
        """
        self._name = value
    
    def rename(self, new_name):
        """
        Renames the worksheet.
        
        Args:
            new_name (str): The new name for the worksheet.
            
        Examples:
            >>> ws.rename("NewSheetName")
        """
        self._name = new_name
    
    @property
    def cells(self):
        """
        Gets the Cells collection for this worksheet.
        
        Returns:
            Cells: The Cells collection containing all cells in the worksheet.
        """
        return self._cells
    
    @property
    def visible(self):
        """
        Gets or sets the visibility state of the worksheet.
        
        Returns:
            bool or str: True for visible, False for hidden, 'veryHidden' for very hidden.
        """
        return self._visible
    
    @visible.setter
    def visible(self, value):
        """
        Sets the visibility state of the worksheet.
        
        Args:
            value (bool or str): True for visible, False for hidden, 'veryHidden' for very hidden.
        """
        self._visible = value
    
    @property
    def tab_color(self):
        """
        Gets or sets the tab color of the worksheet.
        
        Returns:
            str: The tab color in RGB hex format (AARRGGBB), or None if not set.
        """
        return self._tab_color
    
    @tab_color.setter
    def tab_color(self, value):
        """
        Sets the tab color of the worksheet.
        
        Args:
            value (str): The tab color in RGB hex format (AARRGGBB), or None to clear.
        """
        self._tab_color = value
    
    @property
    def auto_filter(self):
        """
        Gets the AutoFilter object for this worksheet.
        
        Returns:
            AutoFilter: The AutoFilter object.
        """
        return self._auto_filter
    
    @property
    def conditional_formats(self):
        """
        Gets the collection of conditional formats for this worksheet.

        Returns:
            ConditionalFormatCollection: The collection of conditional formats.
        """
        return self._conditional_formats

    @property
    def data_validations(self):
        """
        Gets the collection of data validations for this worksheet.

        Data validation allows you to control what data can be entered into cells.
        You can create rules to restrict data entry to specific types, values,
        or ranges, and display messages or errors when invalid data is entered.

        Returns:
            DataValidationCollection: The collection of data validations.

        Examples:
            >>> from aspose_cells import Workbook
            >>> from aspose_cells.data_validation import DataValidationType, DataValidationOperator
            >>>
            >>> wb = Workbook()
            >>> ws = wb.worksheets[0]
            >>>
            >>> # Add a whole number validation
            >>> dv = ws.data_validations.add("A1:A10")
            >>> dv.type = DataValidationType.WHOLE_NUMBER
            >>> dv.operator = DataValidationOperator.BETWEEN
            >>> dv.formula1 = "1"
            >>> dv.formula2 = "100"
            >>> dv.show_error_message = True
            >>> dv.error_message = "Please enter a number between 1 and 100"
            >>>
            >>> # Add a dropdown list validation
            >>> dv2 = ws.data_validations.add("B1:B10")
            >>> dv2.type = DataValidationType.LIST
            >>> dv2.formula1 = '"Red,Green,Blue"'
            >>> dv2.show_dropdown = True
        """
        return self._data_validations

    @property
    def hyperlinks(self):
        """
        Gets the collection of hyperlinks for this worksheet.

        Returns:
            Hyperlinks: The collection of hyperlinks.

        Examples:
            >>> # Add external hyperlink
            >>> link = ws.hyperlinks.add("A1", "https://www.example.com")
            >>> link.text_to_display = "Visit Website"

            >>> # Add internal hyperlink
            >>> link = ws.hyperlinks.add("B2", sub_address="Sheet2!A1")
            >>> link.text_to_display = "Go to Sheet2"
        """
        return self._hyperlinks

    @property
    def protection(self):
        """
        Gets the protection settings for this worksheet.

        Returns:
            SheetProtectionDictWrapper: Dictionary-like object containing protection settings.
        """
        return SheetProtectionDictWrapper(self._properties.protection)
    
    @property
    def page_setup(self):
        """
        Gets the page setup settings for this worksheet.
        
        Returns:
            dict: Dictionary containing page setup settings.
        """
        return self._page_setup
    
    @property
    def page_margins(self):
        """
        Gets the page margins for this worksheet.

        Returns:
            dict: Dictionary containing page margin settings (in inches).
        """
        return self._page_margins

    @property
    def properties(self):
        """
        Gets the worksheet properties.

        Returns:
            WorksheetProperties: The worksheet properties object containing
            view, selection, pane, format, protection, page setup, margins,
            header/footer, and print options.

        Examples:
            >>> ws.properties.view.show_grid_lines = False
            >>> ws.properties.view.zoom_scale = 80
            >>> ws.properties.page_setup.orientation = "landscape"
        """
        return self._properties

    # Methods
    
    def is_protected(self):
        """
        Checks if the worksheet is protected.

        Returns:
            bool: True if the worksheet is protected, False otherwise.
        """
        return self._properties.protection.sheet
    
    def protect(self, password=None, format_cells=None, format_columns=None, format_rows=None,
                insert_columns=None, insert_rows=None, delete_columns=None, delete_rows=None,
                sort=None, auto_filter=None, insert_hyperlinks=None, pivot_tables=None,
                select_locked_cells=None, select_unlocked_cells=None, objects=None, scenarios=None):
        """
        Protects the worksheet with optional password and protection options.

        Args:
            password (str, optional): Password for worksheet protection. Defaults to None.
            format_cells (bool, optional): Allow formatting cells. Defaults to None (keeps current).
            format_columns (bool, optional): Allow formatting columns. Defaults to None (keeps current).
            format_rows (bool, optional): Allow formatting rows. Defaults to None (keeps current).
            insert_columns (bool, optional): Allow inserting columns. Defaults to None (keeps current).
            insert_rows (bool, optional): Allow inserting rows. Defaults to None (keeps current).
            delete_columns (bool, optional): Allow deleting columns. Defaults to None (keeps current).
            delete_rows (bool, optional): Allow deleting rows. Defaults to None (keeps current).
            sort (bool, optional): Allow sorting. Defaults to None (keeps current).
            auto_filter (bool, optional): Allow auto filter. Defaults to None (keeps current).
            insert_hyperlinks (bool, optional): Allow inserting hyperlinks. Defaults to None (keeps current).
            pivot_tables (bool, optional): Allow pivot tables. Defaults to None (keeps current).
            select_locked_cells (bool, optional): Allow selecting locked cells. Defaults to None (keeps current).
            select_unlocked_cells (bool, optional): Allow selecting unlocked cells. Defaults to None (keeps current).
            objects (bool, optional): Protect objects. Defaults to None (keeps current).
            scenarios (bool, optional): Protect scenarios. Defaults to None (keeps current).

        Examples:
            >>> ws.protect()  # Protect without password
            >>> ws.protect("mypassword")  # Protect with password
            >>> ws.protect("secure", format_cells=True, sort=False)  # Protect with options
        """
        # Enable sheet protection
        self._properties.protection.sheet = True

        # Set password (will be hashed when saving to XML)
        if password is not None:
            self._properties.protection.password = password

        # Set protection options if provided
        if format_cells is not None:
            self._properties.protection.format_cells = format_cells
        if format_columns is not None:
            self._properties.protection.format_columns = format_columns
        if format_rows is not None:
            self._properties.protection.format_rows = format_rows
        if insert_columns is not None:
            self._properties.protection.insert_columns = insert_columns
        if insert_rows is not None:
            self._properties.protection.insert_rows = insert_rows
        if delete_columns is not None:
            self._properties.protection.delete_columns = delete_columns
        if delete_rows is not None:
            self._properties.protection.delete_rows = delete_rows
        if sort is not None:
            self._properties.protection.sort = sort
        if auto_filter is not None:
            self._properties.protection.auto_filter = auto_filter
        if insert_hyperlinks is not None:
            self._properties.protection.insert_hyperlinks = insert_hyperlinks
        if pivot_tables is not None:
            self._properties.protection.pivot_tables = pivot_tables
        if select_locked_cells is not None:
            self._properties.protection.select_locked_cells = select_locked_cells
        if select_unlocked_cells is not None:
            self._properties.protection.select_unlocked_cells = select_unlocked_cells
        if objects is not None:
            self._properties.protection.objects = objects
        if scenarios is not None:
            self._properties.protection.scenarios = scenarios
    
    def unprotect(self, password=None):
        """
        Unprotects the worksheet.

        Args:
            password (str, optional): Password to unprotect the worksheet. Defaults to None.
                Note: Password validation is not currently implemented.

        Examples:
            >>> ws.unprotect()  # Unprotect without password
            >>> ws.unprotect("mypassword")  # Unprotect with password
        """
        # Note: We don't validate the password here - just disable protection
        # In a full implementation, we would check if the password matches
        self._properties.protection.sheet = False
        self._properties.protection.password = None
    
    def set_view(self, zoom=None, show_grid_lines=None, show_row_col_headers=None):
        """
        Sets view options for the worksheet.
        
        Args:
            zoom (int, optional): Zoom percentage (10-400). Defaults to None.
            show_grid_lines (bool, optional): Whether to show grid lines. Defaults to None.
            show_row_col_headers (bool, optional): Whether to show row and column headers. Defaults to None.
        """
        # This is a placeholder for future implementation
        pass
    
    def copy(self, name=None):
        """
        Creates a copy of the worksheet.
        
        Args:
            name (str, optional): Name for the copied worksheet. If None, a default name is generated.
            
        Returns:
            Worksheet: The copied worksheet.
        """
        new_ws = Worksheet(name if name else f"{self._name} (copy)")
        # Copy cells
        for ref, cell in self._cells._cells.items():
            new_ws._cells._cells[ref] = Cell(cell.value, cell.formula)
            if cell.style:
                new_ws._cells._cells[ref].style = cell.style.copy()
        return new_ws
    
    def delete(self):
        """
        Deletes the worksheet from the workbook.
        
        Note: This method is a placeholder. The actual deletion is handled by the Workbook class.
        """
        pass
    
    def move(self, index):
        """
        Moves the worksheet to the specified position in the workbook.
        
        Args:
            index (int): Zero-based index where to move the worksheet.
            
        Note: This method is a placeholder. The actual move is handled by the Workbook class.
        """
        pass
    
    def select(self):
        """
        Selects the worksheet.
        
        Note: This method is a placeholder. The actual selection is handled by the Workbook class.
        """
        pass
    
    def activate(self):
        """
        Activates the worksheet (makes it the active worksheet).
        
        Note: This method is a placeholder. The actual activation is handled by the Workbook class.
        """
        pass
    
    def calculate_formula(self):
        """
        Calculates formulas in the worksheet.
        
        Note: This method is a placeholder for future implementation.
        """
        pass
    
    def get_range(self, start_cell, end_cell=None):
        """
        Gets a range of cells.
        
        Args:
            start_cell (str): The starting cell reference (e.g., "A1").
            end_cell (str, optional): The ending cell reference. If None, returns a single cell.
            
        Returns:
            The range or cell.
            
        Note: This method is a placeholder for future implementation.
        """
        pass
