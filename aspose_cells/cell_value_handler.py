"""
Aspose.Cells for Python - Cell Value Handler Module

This module provides comprehensive cell value import/export functionality according to ECMA-376 specification.
It handles all cell value data types: shared strings, inline strings, numeric, boolean, error, and formula.

ECMA-376 Reference:
- Part 1, Section 18.3.1.4 - cell (Cell)
- Part 1, Section 18.3.1.16 - f (Formula)
- Part 1, Section 18.3.1.96 - v (Cell Value)
- Part 1, Section 18.3.1.53 - is (Rich Text Inline)
"""

from datetime import datetime, date, time


class CellValueHandler:
    """
    Handles cell value import and export operations according to ECMA-376 specification.
    
    This class provides methods to:
    - Import cell values from XML elements
    - Export cell values to XML elements
    - Determine cell type attributes
    - Format values for XML output
    """
    
    # ECMA-376 Cell Type Constants
    TYPE_SHARED_STRING = 's'      # Shared string index
    TYPE_INLINE_STRING = 'str'     # Inline string
    TYPE_NUMBER = 'n'             # Number (default)
    TYPE_BOOLEAN = 'b'             # Boolean
    TYPE_ERROR = 'e'              # Error
    
    # ECMA-376 Error Values
    ERROR_NULL = '#NULL!'
    ERROR_DIV_0 = '#DIV/0!'
    ERROR_VALUE = '#VALUE!'
    ERROR_REF = '#REF!'
    ERROR_NAME = '#NAME?'
    ERROR_NUM = '#NUM!'
    ERROR_NA = '#N/A'
    
    # Valid error values according to ECMA-376
    VALID_ERRORS = {
        ERROR_NULL, ERROR_DIV_0, ERROR_VALUE, ERROR_REF,
        ERROR_NAME, ERROR_NUM, ERROR_NA
    }
    
    @staticmethod
    def get_cell_type(value):
        """
        Determines the ECMA-376 cell type attribute for a given value.
        
        Args:
            value: The cell value (can be None, int, float, str, bool, datetime, date, time)
            
        Returns:
            str: The ECMA-376 cell type attribute ('s', 'str', 'n', 'b', 'e', or None for default)
            
        Examples:
            >>> CellValueHandler.get_cell_type("Hello")
            's'
            >>> CellValueHandler.get_cell_type(42)
            None  # Defaults to 'n' (number)
            >>> CellValueHandler.get_cell_type(True)
            'b'
            >>> CellValueHandler.get_cell_type("#N/A")
            'e'
        """
        if value is None:
            return None
        
        # Check for error values first (strings that match error patterns)
        if isinstance(value, str) and value in CellValueHandler.VALID_ERRORS:
            return CellValueHandler.TYPE_ERROR
        
        # Boolean values
        if isinstance(value, bool):
            return CellValueHandler.TYPE_BOOLEAN
        
        # Numeric values (default type, can be omitted)
        if isinstance(value, (int, float)):
            return None  # Default to 'n'
        
        # Date/time values (stored as numbers in Excel)
        if isinstance(value, (datetime, date, time)):
            return None  # Stored as numbers
        
        # String values
        if isinstance(value, str):
            # Check if string represents a boolean
            if value.upper() in ('TRUE', 'FALSE'):
                return CellValueHandler.TYPE_BOOLEAN
            # Check if string represents an error
            if value.upper() in (e.upper() for e in CellValueHandler.VALID_ERRORS):
                return CellValueHandler.TYPE_ERROR
            # Regular string - use shared string type
            return CellValueHandler.TYPE_SHARED_STRING
        
        return None
    
    @staticmethod
    def format_value_for_xml(value, cell_type=None):
        """
        Formats a cell value for XML output according to ECMA-376 specification.
        
        Args:
            value: The cell value to format
            cell_type (str, optional): The cell type attribute. If None, will be determined automatically.
            
        Returns:
            tuple: (formatted_value_str, cell_type) - The formatted value string and cell type
            
        Examples:
            >>> CellValueHandler.format_value_for_xml("Hello")
            ('0', 's')  # Index into shared strings table
            >>> CellValueHandler.format_value_for_xml(42.5)
            ('42.5', None)
            >>> CellValueHandler.format_value_for_xml(True)
            ('1', 'b')
            >>> CellValueHandler.format_value_for_xml("#N/A")
            ('#N/A', 'e')
        """
        if value is None:
            return (None, None)
        
        # Determine cell type if not provided
        if cell_type is None:
            cell_type = CellValueHandler.get_cell_type(value)
        
        # Format based on type
        if cell_type == CellValueHandler.TYPE_BOOLEAN:
            # Boolean: 1 for True, 0 for False
            return ('1' if value else '0', cell_type)
        
        elif cell_type == CellValueHandler.TYPE_ERROR:
            # Error: use the error string as-is
            return (value, cell_type)
        
        elif cell_type == CellValueHandler.TYPE_SHARED_STRING:
            # Shared string: return the string itself (index will be assigned later)
            return (value, cell_type)
        
        elif cell_type == CellValueHandler.TYPE_INLINE_STRING:
            # Inline string: return the string as-is
            return (value, cell_type)
        
        else:
            # Number (default type)
            if isinstance(value, (datetime, date, time)):
                # Convert date/time to Excel serial date number
                serial_date = CellValueHandler._datetime_to_excel_serial(value)
                return (str(serial_date), None)
            else:
                # Regular number
                return (str(value), None)
    
    @staticmethod
    def parse_value_from_xml(value_str, cell_type, shared_strings=None):
        """
        Parses a cell value from XML according to ECMA-376 specification.
        
        Args:
            value_str (str): The value string from XML <v> element
            cell_type (str): The cell type attribute ('s', 'str', 'n', 'b', 'e', or None)
            shared_strings (list, optional): List of shared strings for type 's'
            
        Returns:
            The parsed value (int, float, str, bool, or None)
            
        Examples:
            >>> CellValueHandler.parse_value_from_xml("42", None)
            42
            >>> CellValueHandler.parse_value_from_xml("0", "b")
            False
            >>> CellValueHandler.parse_value_from_xml("5", "s", ["Hello", "World"])
            'World'
            >>> CellValueHandler.parse_value_from_xml("#N/A", "e")
            '#N/A'
        """
        if value_str is None or value_str == '':
            return None
        
        # Default to number type if not specified
        if cell_type is None:
            cell_type = CellValueHandler.TYPE_NUMBER
        
        # Parse based on type
        if cell_type == CellValueHandler.TYPE_SHARED_STRING:
            # Shared string: value_str is an index into shared strings table
            if shared_strings is not None:
                try:
                    index = int(value_str)
                    if 0 <= index < len(shared_strings):
                        return shared_strings[index]
                except (ValueError, IndexError):
                    pass
            return value_str
        
        elif cell_type == CellValueHandler.TYPE_INLINE_STRING:
            # Inline string: value_str is the actual string
            return value_str
        
        elif cell_type == CellValueHandler.TYPE_BOOLEAN:
            # Boolean: 1 for True, 0 for False
            return bool(int(value_str))
        
        elif cell_type == CellValueHandler.TYPE_ERROR:
            # Error: value_str is the error string
            return value_str
        
        else:
            # Number (default type)
            try:
                if '.' in value_str or 'e' in value_str.lower():
                    return float(value_str)
                else:
                    return int(value_str)
            except ValueError:
                return value_str
    
    @staticmethod
    def _datetime_to_excel_serial(dt):
        """
        Converts a datetime/date/time object to Excel serial date number.
        
        Excel stores dates as serial numbers where:
        - 1 = January 1, 1900 (incorrectly treats 1900 as a leap year)
        - 2 = January 2, 1900
        - Fractional part represents time of day
        
        Args:
            dt: datetime, date, or time object
            
        Returns:
            float: Excel serial date number
        """
        # Excel base date (January 1, 1900)
        excel_epoch = datetime(1899, 12, 30)  # Adjusted for Excel's 1900 leap year bug
        
        if isinstance(dt, time):
            # Time only: convert to datetime with base date
            dt = datetime.combine(datetime.today().date(), dt)
        
        if isinstance(dt, date) and not isinstance(dt, datetime):
            # Date only: convert to datetime at midnight
            dt = datetime.combine(dt, time.min)
        
        # Calculate difference in days
        delta = dt - excel_epoch
        serial_date = delta.days + delta.seconds / 86400.0
        
        return serial_date
    
    @staticmethod
    def excel_serial_to_datetime(serial_date):
        """
        Converts an Excel serial date number to a datetime object.
        
        Args:
            serial_date (float): Excel serial date number
            
        Returns:
            datetime: The corresponding datetime object
        """
        # Excel base date (January 1, 1900)
        excel_epoch = datetime(1899, 12, 30)  # Adjusted for Excel's 1900 leap year bug
        
        # Calculate datetime from serial
        days = int(serial_date)
        fraction = serial_date - days
        seconds = int(fraction * 86400)
        
        dt = excel_epoch + timedelta(days=days, seconds=seconds)
        return dt
    
    @staticmethod
    def is_error_value(value):
        """
        Checks if a value is a valid ECMA-376 error value.
        
        Args:
            value: The value to check
            
        Returns:
            bool: True if the value is a valid error value, False otherwise
        """
        return isinstance(value, str) and value in CellValueHandler.VALID_ERRORS
    
    @staticmethod
    def get_error_type(value):
        """
        Returns the error type name for a given error value.
        
        Args:
            value: The error value string
            
        Returns:
            str: The error type name or None if not a valid error
        """
        if not CellValueHandler.is_error_value(value):
            return None
        
        error_map = {
            CellValueHandler.ERROR_NULL: 'NULL',
            CellValueHandler.ERROR_DIV_0: 'DIV_0',
            CellValueHandler.ERROR_VALUE: 'VALUE',
            CellValueHandler.ERROR_REF: 'REF',
            CellValueHandler.ERROR_NAME: 'NAME',
            CellValueHandler.ERROR_NUM: 'NUM',
            CellValueHandler.ERROR_NA: 'NA'
        }
        
        return error_map.get(value)


# Import timedelta for datetime conversion
from datetime import timedelta
