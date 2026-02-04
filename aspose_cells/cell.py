"""
Aspose.Cells for Python - Cell Module

This module provides the Cell class which represents a single cell in a worksheet.
The Cell class provides methods and properties for accessing and modifying cell values,
formulas, styles, and comments.

Compatible with Aspose.Cells for .NET API structure.
"""

import sys
from datetime import datetime, date, time
from .style import Style


class Cell:
    """
    Represents a single cell in a worksheet.
    
    The Cell class provides properties and methods for working with individual cells,
    including value access, formula handling, style application, and comment management.
    
    Examples:
        >>> from aspose_cells import Workbook
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> cell = ws.cells['A1']
        >>> cell.value = "Hello"
        >>> cell.style.font.bold = True
        >>> cell.set_comment("This is a note", "Author")
    """
    
    def __init__(self, value=None, formula=None):
        """
        Initializes a new instance of the Cell class.
        
        Args:
            value: The value to store in the cell. Can be None, int, float, str, bool,
                   datetime, date, or time objects.
            formula (str, optional): The formula to store in the cell.
            
        Examples:
            >>> cell = Cell()
            >>> cell = Cell("Hello")
            >>> cell = Cell(42, "=SUM(A1:B1)")
        """
        self._value = value
        self._formula = formula
        self._style = Style()
        self._comment = None
        self._style_index = 0  # Internal use for saving
        
        # Debug logging
        if '--debug' in sys.argv:
            print(f"DEBUG Cell.__init__: Created new cell with value={value}, formula={formula}")
            print(f"  Initial borders: top={self._style.borders.top.line_style}, {self._style.borders.top.color}")
    
    # Properties
    
    @property
    def value(self):
        """
        Gets or sets the value of the cell.
        
        Returns:
            The cell value. Can be None, int, float, str, bool, datetime, date, or time.
            
        Examples:
            >>> cell.value = "Hello"
            >>> print(cell.value)
        """
        return self._value
    
    @value.setter
    def value(self, val):
        """
        Sets the value of the cell.
        
        Args:
            val: The value to set. Can be None, int, float, str, bool, datetime, date, or time.
        """
        self._value = val
    
    @property
    def formula(self):
        """
        Gets or sets the formula of the cell.
        
        Returns:
            str or None: The formula string, or None if no formula is set.
            
        Examples:
            >>> cell.formula = "=SUM(A1:B1)"
            >>> print(cell.formula)
        """
        return self._formula
    
    @formula.setter
    def formula(self, val):
        """
        Sets the formula of the cell.
        
        Args:
            val (str): The formula string to set.
        """
        self._formula = val
    
    @property
    def style(self):
        """
        Gets or sets the style of the cell.
        
        Returns:
            Style: The Style object containing cell formatting properties.
            
        Examples:
            >>> cell.style.font.bold = True
            >>> cell.style.set_fill_color('FFFF0000')
        """
        return self._style
    
    @style.setter
    def style(self, value):
        """
        Sets the style of the cell.
        
        Args:
            value (Style): The Style object to apply to the cell.
        """
        if '--debug' in sys.argv:
            print(f"DEBUG Cell.__setattr__: Setting style to {value}")
            if hasattr(value, 'borders'):
                print(f"  New style borders: top={value.borders.top.line_style}, {value.borders.top.color}")
        object.__setattr__(self, '_style', value)
    
    @property
    def comment(self):
        """
        Gets the comment associated with the cell.
        
        Returns:
            dict or None: Dictionary containing 'text' and 'author' keys, or None if no comment.
            
        Examples:
            >>> if cell.comment:
            ...     print(cell.comment['text'])
        """
        return self._comment
    
    # Data type detection
    
    @property
    def data_type(self):
        """
        Gets the data type of the cell value.
        
        Returns:
            str: The data type of the cell value. Possible values:
                - 'none': Cell is empty (value is None)
                - 'boolean': Boolean value (True/False or string "TRUE"/"FALSE")
                - 'numeric': Integer or floating-point number
                - 'datetime': datetime, date, or time object
                - 'string': Text string
                - 'unknown': Any other type
                
        Examples:
            >>> cell.value = "Hello"
            >>> print(cell.data_type)  # 'string'
            >>> cell.value = 42
            >>> print(cell.data_type)  # 'numeric'
        """
        if self._value is None:
            return 'none'
        elif isinstance(self._value, bool):
            return 'boolean'
        elif isinstance(self._value, (int, float)):
            return 'numeric'
        elif isinstance(self._value, (datetime, date, time)):
            return 'datetime'
        elif isinstance(self._value, str):
            # Check for boolean strings
            if self._value.upper() in ('TRUE', 'FALSE'):
                return 'boolean'
            return 'string'
        else:
            return 'unknown'
    
    # Cell value methods
    
    def is_empty(self):
        """
        Checks if the cell is empty.
        
        Returns:
            bool: True if the cell value is None, False otherwise.
            
        Examples:
            >>> if cell.is_empty():
            ...     print("Cell is empty")
        """
        return self._value is None
    
    def clear_value(self):
        """
        Clears the value of the cell (sets it to None).
        
        Examples:
            >>> cell.clear_value()
        """
        self._value = None
    
    def clear_formula(self):
        """
        Clears the formula of the cell (sets it to None).
        
        Examples:
            >>> cell.clear_formula()
        """
        self._formula = None
    
    def clear(self):
        """
        Clears both the value and formula of the cell.
        
        Examples:
            >>> cell.clear()
        """
        self._value = None
        self._formula = None
    
    # Comment methods
    
    def set_comment(self, text, author='None', width=None, height=None):
        """
        Sets a comment on the cell.

        Args:
            text (str): The comment text.
            author (str, optional): The author of the comment. Defaults to 'None'.
            width (float, optional): The width of the comment box in points. Defaults to None (uses Excel default).
            height (float, optional): The height of the comment box in points. Defaults to None (uses Excel default).

        Examples:
            >>> cell.set_comment("This is important", "John")
            >>> cell.set_comment("Note")  # Author defaults to 'None'
            >>> cell.set_comment("Large note", "John", width=200, height=100)
        """
        # If author is empty string, set it to "None"
        if author == '':
            author = 'None'
        self._comment = {
            'text': text,
            'author': author,
            'width': width,
            'height': height
        }
    
    def get_comment(self):
        """
        Gets the comment from the cell.
        
        Returns:
            dict or None: Dictionary containing 'text' and 'author' keys, or None if no comment.
            
        Examples:
            >>> comment = cell.get_comment()
            >>> if comment:
            ...     print(f"{comment['author']}: {comment['text']}")
        """
        return self._comment
    
    def clear_comment(self):
        """
        Clears the comment from the cell.
        
        Examples:
            >>> cell.clear_comment()
        """
        self._comment = None
    
    def has_comment(self):
        """
        Checks if the cell has a comment.

        Returns:
            bool: True if the cell has a comment, False otherwise.

        Examples:
            >>> if cell.has_comment():
            ...     print("Cell has a comment")
        """
        return self._comment is not None

    def set_comment_size(self, width, height):
        """
        Sets the size of the comment box.

        Args:
            width (float): The width of the comment box in points.
            height (float): The height of the comment box in points.

        Examples:
            >>> cell.set_comment("Note", "John")
            >>> cell.set_comment_size(150, 80)
        """
        if self._comment is None:
            raise ValueError("Cell has no comment. Call set_comment() first.")
        self._comment['width'] = width
        self._comment['height'] = height

    def get_comment_size(self):
        """
        Gets the size of the comment box.

        Returns:
            tuple or None: A tuple of (width, height) in points, or None if no size is set.

        Examples:
            >>> size = cell.get_comment_size()
            >>> if size:
            ...     print(f"Width: {size[0]}, Height: {size[1]}")
        """
        if self._comment is None:
            return None
        width = self._comment.get('width')
        height = self._comment.get('height')
        if width is not None and height is not None:
            return (width, height)
        return None
    
    # Style methods
    
    def apply_style(self, style):
        """
        Applies a style to the cell.
        
        Args:
            style (Style): The Style object to apply to the cell.
            
        Examples:
            >>> from aspose_cells import Style
            >>> style = Style()
            >>> style.font.bold = True
            >>> cell.apply_style(style)
        """
        self._style = style
    
    def get_style(self):
        """
        Gets the style of the cell.
        
        Returns:
            Style: The Style object containing cell formatting properties.
            
        Examples:
            >>> style = cell.get_style()
            >>> print(style.font.name)
        """
        return self._style
    
    def clear_style(self):
        """
        Clears the style of the cell (resets to default).
        
        Examples:
            >>> cell.clear_style()
        """
        self._style = Style()
    
    # Formula methods
    
    def has_formula(self):
        """
        Checks if the cell has a formula.
        
        Returns:
            bool: True if the cell has a formula, False otherwise.
            
        Examples:
            >>> if cell.has_formula():
            ...     print("Cell contains a formula")
        """
        return self._formula is not None
    
    def is_numeric_value(self):
        """
        Checks if the cell value is numeric.
        
        Returns:
            bool: True if the cell value is an int or float, False otherwise.
            
        Examples:
            >>> if cell.is_numeric_value():
            ...     print("Cell contains a number")
        """
        return isinstance(self._value, (int, float))
    
    def is_text_value(self):
        """
        Checks if the cell value is text.
        
        Returns:
            bool: True if the cell value is a string, False otherwise.
            
        Examples:
            >>> if cell.is_text_value():
            ...     print("Cell contains text")
        """
        return isinstance(self._value, str)
    
    def is_boolean_value(self):
        """
        Checks if the cell value is boolean.
        
        Returns:
            bool: True if the cell value is a bool or boolean string ("TRUE"/"FALSE"), False otherwise.
            
        Examples:
            >>> if cell.is_boolean_value():
            ...     print("Cell contains a boolean value")
        """
        return self.data_type == 'boolean'
    
    def is_date_time_value(self):
        """
        Checks if the cell value is a date/time.
        
        Returns:
            bool: True if the cell value is a datetime, date, or time object, False otherwise.
            
        Examples:
            >>> if cell.is_date_time_value():
            ...     print("Cell contains a date/time value")
        """
        return self.data_type == 'datetime'
    
    # String representation
    
    def __str__(self):
        """
        Returns a string representation of the cell value.
        
        Returns:
            str: String representation of the cell value, or empty string if value is None.
        """
        return str(self._value) if self._value is not None else ''
    
    def __repr__(self):
        """
        Returns a detailed string representation of the cell.
        
        Returns:
            str: Detailed representation including value, formula, and data type.
        """
        return f"Cell(value={self._value!r}, formula={self._formula!r}, type={self.data_type})"
    
    # Comparison methods
    
    def __eq__(self, other):
        """
        Checks if two cells are equal based on their values.
        
        Args:
            other: Another Cell object or value to compare with.
            
        Returns:
            bool: True if values are equal, False otherwise.
        """
        if isinstance(other, Cell):
            return self._value == other._value
        return self._value == other
    
    def __ne__(self, other):
        """
        Checks if two cells are not equal based on their values.
        
        Args:
            other: Another Cell object or value to compare with.
            
        Returns:
            bool: True if values are not equal, False otherwise.
        """
        return not self.__eq__(other)
