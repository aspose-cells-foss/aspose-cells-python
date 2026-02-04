"""
Aspose.Cells for Python - Style Module

This module provides classes for cell styling including Font, Fill, Border, Borders,
Alignment, NumberFormat, Protection, and Style. These classes represent the
formatting properties that can be applied to cells in an Excel worksheet.

Compatible with Aspose.Cells for .NET API structure.
"""


class Font:
    """
    Represents font settings for a cell or range of cells.
    
    The Font class provides properties for controlling the appearance of text,
    including font name, size, color, and various text effects.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.font.name = "Arial"
        >>> style.font.size = 12
        >>> style.font.bold = True
        >>> style.font.color = "FFFF0000"  # Red color
    """
    
    def __init__(self, name='Calibri', size=11, color='FF000000', bold=False, italic=False, underline=False, strikethrough=False):
        """
        Initializes a new instance of the Font class.
        
        Args:
            name (str, optional): Font name. Defaults to 'Calibri'.
            size (int, optional): Font size in points. Defaults to 11.
            color (str, optional): Font color in AARRGGBB hex format. Defaults to 'FF000000' (black).
            bold (bool, optional): Whether text is bold. Defaults to False.
            italic (bool, optional): Whether text is italic. Defaults to False.
            underline (bool, optional): Whether text is underlined. Defaults to False.
            strikethrough (bool, optional): Whether text has strikethrough. Defaults to False.
            
        Examples:
            >>> font = Font()
            >>> font = Font(name="Arial", size=12, bold=True)
        """
        self.name = name
        self.size = size
        self.color = color
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strikethrough = strikethrough

class Fill:
    """
    Represents fill settings for a cell or range of cells.
    
    The Fill class provides properties for controlling the background appearance of cells,
    including solid colors, patterns, and gradients.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.fill.set_solid_fill('FFFF0000')  # Red background
        >>> style.fill.set_pattern_fill('gray125', 'FFCCCCCC', 'FFFFFFFF')
    """
    
    def __init__(self, pattern_type='none', foreground_color='FFFFFFFF', background_color='FFFFFFFF'):
        """
        Initializes a new instance of Fill class.
        
        Args:
            pattern_type (str, optional): Fill pattern type. Common values: 'none', 'solid', 'gray125'.
                                       Defaults to 'none'.
            foreground_color (str, optional): Foreground color in AARRGGBB hex format.
                                            Defaults to 'FFFFFFFF' (white).
            background_color (str, optional): Background color in AARRGGBB hex format.
                                            Defaults to 'FFFFFFFF' (white).
                                            
        Examples:
            >>> fill = Fill()
            >>> fill = Fill(pattern_type='solid', foreground_color='FFFF0000')
        """
        self.pattern_type = pattern_type
        self.foreground_color = foreground_color
        self.background_color = background_color
    
    def set_solid_fill(self, color):
        """
        Sets a solid fill pattern with specified color.
        
        Args:
            color (str): Fill color in AARRGGBB hex format.
            
        Examples:
            >>> fill.set_solid_fill('FFFF0000')  # Red solid fill
        """
        self.pattern_type = 'solid'
        self.foreground_color = color
        self.background_color = color
    
    def set_gradient_fill(self, start_color, end_color):
        """
        Sets a gradient fill pattern (currently simplified as solid fill).
        
        Args:
            start_color (str): Start color in AARRGGBB hex format.
            end_color (str): End color in AARRGGBB hex format.
            
        Examples:
            >>> fill.set_gradient_fill('FF0000FF', 'FFFFFFFF')  # Blue to white gradient
        """
        self.pattern_type = 'solid'
        self.foreground_color = start_color
        self.background_color = end_color
    
    def set_pattern_fill(self, pattern_type, fg_color='FFFFFFFF', bg_color='FFFFFFFF'):
        """
        Sets a pattern fill with specified pattern type and colors.
        
        Args:
            pattern_type (str): Pattern type. Common values: 'none', 'solid', 'gray125', 'gray0625',
                              'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid',
                              'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
                              'lightGrid', 'lightTrellis', 'mediumGray', 'darkGray'.
            fg_color (str, optional): Foreground color in AARRGGBB hex format. Defaults to 'FFFFFFFF'.
            bg_color (str, optional): Background color in AARRGGBB hex format. Defaults to 'FFFFFFFF'.
            
        Examples:
            >>> fill.set_pattern_fill('gray125')
            >>> fill.set_pattern_fill('lightGrid', 'FFCCCCCC', 'FFFFFFFF')
        """
        self.pattern_type = pattern_type
        self.foreground_color = fg_color
        self.background_color = bg_color
    
    def set_no_fill(self):
        """
        Sets no fill (transparent background).
        
        Examples:
            >>> fill.set_no_fill()
        """
        self.pattern_type = 'none'
        self.foreground_color = 'FFFFFFFF'
        self.background_color = 'FFFFFFFF'

class Border:
    """
    Represents border settings for a single side of a cell or range of cells.
    
    The Border class provides properties for controlling the appearance of cell borders,
    including line style, color, and weight.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.borders.top.line_style = 'thin'
        >>> style.borders.top.color = 'FFFF0000'  # Red border
    """
    
    def __init__(self, line_style='none', color='FF000000', weight=1):
        """
        Initializes a new instance of Border class.
        
        Args:
            line_style (str, optional): Border line style. Common values: 'none', 'thin', 'medium', 'thick',
                                      'dotted', 'dashed', 'double', 'hair', 'mediumDashed',
                                      'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot',
                                      'slantDashDot'. Defaults to 'none'.
            color (str, optional): Border color in AARRGGBB hex format. Defaults to 'FF000000' (black).
            weight (int, optional): Border weight/thickness. Defaults to 1.
            
        Examples:
            >>> border = Border()
            >>> border = Border(line_style='thin', color='FFFF0000', weight=1)
        """
        self.line_style = line_style
        self.color = color
        self.weight = weight

class Borders:
    """
    Represents border settings for all sides of a cell or range of cells.
    
    The Borders class provides properties for controlling borders on all four sides
    of a cell, as well as diagonal borders.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.borders.set_border('all', 'thin', 'FF000000')  # Thin black border on all sides
        >>> style.borders.set_border('top', 'thick', 'FFFF0000')  # Thick red top border
    """
    
    def __init__(self):
        """
        Initializes a new instance of Borders class.
        
        Creates Border objects for top, bottom, left, right, and diagonal sides.
        
        Examples:
            >>> borders = Borders()
        """
        self.top = Border()
        self.bottom = Border()
        self.left = Border()
        self.right = Border()
        # Diagonal borders
        self.diagonal_up = False
        self.diagonal_down = False
        self.diagonal = Border()  # Style for diagonal lines

class Alignment:
    """
    Represents alignment settings for a cell or range of cells.
    
    The Alignment class provides properties for controlling text alignment within cells,
    including horizontal and vertical alignment, text wrapping, rotation, and indentation.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.alignment.horizontal = 'center'
        >>> style.alignment.vertical = 'center'
        >>> style.alignment.wrap_text = True
    """
    
    def __init__(self, horizontal='general', vertical='bottom', wrap_text=False, indent=0,
                 text_rotation=0, shrink_to_fit=False, reading_order=0, relative_indent=0):
        """
        Initializes a new instance of Alignment class.
        
        Args:
            horizontal (str, optional): Horizontal alignment. Valid values: 'general', 'left', 'center', 'right',
                                      'fill', 'justify', 'centerContinuous', 'distributed'. Defaults to 'general'.
            vertical (str, optional): Vertical alignment. Valid values: 'top', 'center', 'bottom', 'justify',
                                    'distributed'. Defaults to 'bottom'.
            wrap_text (bool, optional): Whether text wraps within the cell. Defaults to False.
            indent (int, optional): Indent level (0-250). Defaults to 0.
            text_rotation (int, optional): Text rotation in degrees (0-180, or 255 for vertical). Defaults to 0.
            shrink_to_fit (bool, optional): Whether text shrinks to fit cell width. Defaults to False.
            reading_order (int, optional): Reading order. 0=Context, 1=Left-to-Right, 2=Right-to-Left.
                                        Defaults to 0.
            relative_indent (int, optional): Relative indent. Defaults to 0.
            
        Examples:
            >>> alignment = Alignment()
            >>> alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        """
        self.horizontal = horizontal
        self.vertical = vertical
        self.wrap_text = wrap_text
        self.indent = indent
        self.text_rotation = text_rotation  # 0-180 degrees
        self.shrink_to_fit = shrink_to_fit
        self.reading_order = reading_order  # 0=Context, 1=Left-to-Right, 2=Right-to-Left
        self.relative_indent = relative_indent

class NumberFormat:
    """
    Represents number format settings for a cell or range of cells.
    
    The NumberFormat class provides methods for applying built-in and custom number formats
    to display values in various formats such as currency, percentage, date, time, etc.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.number_format = '0.00'  # Two decimal places
        >>> style.number_format = '#,##0.00'  # Thousands separator with two decimals
        >>> style.number_format = '0%'  # Percentage format
    """
    
    # Built-in number formats (compatible with Excel)
    BUILTIN_FORMATS = {
        0: 'General',
        1: '0',
        2: '0.00',
        3: '#,##0',
        4: '#,##0.00',
        5: '$#,##0_);($#,##0)',
        6: '$#,##0_);[Red]($#,##0)',
        7: '$#,##0.00_);($#,##0.00)',
        8: '$#,##0.00_);[Red]($#,##0.00)',
        9: '0%',
        10: '0.00%',
        11: '0.00E+00',
        12: '# ?/?',
        13: '# ??/??',
        14: 'mm-dd-yy',
        15: 'd-mmm-yy',
        16: 'd-mmm',
        17: 'mmm-yy',
        18: 'h:mm AM/PM',
        19: 'h:mm:ss AM/PM',
        20: 'h:mm',
        21: 'h:mm:ss',
        22: 'm/d/yy h:mm',
        37: '#,##0_);(#,##0)',
        38: '#,##0_);[Red](#,##0)',
        39: '#,##0.00_);(#,##0.00)',
        40: '#,##0.00_);[Red](#,##0.00)',
        41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
        42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
        43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
        44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
        45: 'mm:ss',
        46: '[h]:mm:ss',
        47: 'mm:ss.0',
        48: '##0.0E+0',
        49: '@',
        # Additional common formats
        50: 'General',
        51: '0_);(0)',
        52: '0_);[Red](0)',
        53: '0_);(0)',
        54: '0_);[Red](0)',
        55: '0_);(0)',
        56: '0_);[Red](0)',
        57: '0_);(0)',
        58: '0_);[Red](0)',
        59: '0_);(0)',
        60: '0_);[Red](0)',
        61: '0_);(0)',
        62: '0_);[Red](0)',
        63: '0_);(0)',
        64: '0_);[Red](0)',
        65: '0_);(0)',
        66: '0_);[Red](0)',
        67: '0_);(0)',
        68: '0_);[Red](0)',
        69: '0_);(0)',
        70: '0_);[Red](0)',
        71: '0_);(0)',
        72: '0_);[Red](0)',
        73: '0_);(0)',
        74: '0_);[Red](0)',
        75: '0_);(0)',
        76: '0_);[Red](0)',
        77: '0_);(0)',
        78: '0_);[Red](0)',
        79: '0_);(0)',
        80: '0_);[Red](0)',
        81: '0_);(0)',
        82: '0_);[Red](0)',
    }
    
    @staticmethod
    def get_builtin_format(format_id):
        """
        Gets a built-in format string by format ID.
        
        Args:
            format_id (int): The built-in format ID (0-164).
            
        Returns:
            str: The format string, or 'General' if the ID is not found.
            
        Examples:
            >>> NumberFormat.get_builtin_format(2)  # Returns '0.00'
            >>> NumberFormat.get_builtin_format(14)  # Returns 'mm-dd-yy'
        """
        return NumberFormat.BUILTIN_FORMATS.get(format_id, 'General')
    
    @staticmethod
    def is_builtin_format(format_code):
        """
        Checks if a format code is a built-in format.
        
        Args:
            format_code (str): The format code to check.
            
        Returns:
            bool: True if the format code is built-in, False otherwise.
            
        Examples:
            >>> NumberFormat.is_builtin_format('0.00')  # Returns True
            >>> NumberFormat.is_builtin_format('Custom')  # Returns False
        """
        return format_code in NumberFormat.BUILTIN_FORMATS.values()
    
    @staticmethod
    def lookup_builtin_format(format_code):
        """
        Looks up the format ID for a built-in format code.
        
        Args:
            format_code (str): The format code to look up.
            
        Returns:
            int or None: The format ID if found, None otherwise.
            
        Examples:
            >>> NumberFormat.lookup_builtin_format('0.00')  # Returns 2
            >>> NumberFormat.lookup_builtin_format('mm-dd-yy')  # Returns 14
        """
        for fmt_id, fmt_code in NumberFormat.BUILTIN_FORMATS.items():
            if fmt_code == format_code:
                return fmt_id
        return None

class Protection:
    """
    Represents protection settings for a cell or range of cells.
    
    The Protection class provides properties for controlling cell protection
    when the worksheet is protected.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.protection.locked = False  # Cell can be edited when sheet is protected
        >>> style.protection.hidden = True  # Formula is hidden when sheet is protected
    """
    
    def __init__(self, locked=True, hidden=False):
        """
        Initializes a new instance of Protection class.
        
        Args:
            locked (bool, optional): Whether the cell is locked when the worksheet is protected.
                                   Defaults to True.
            hidden (bool, optional): Whether the cell's formula is hidden when the worksheet is protected.
                                   Defaults to False.
                                   
        Examples:
            >>> protection = Protection()
            >>> protection = Protection(locked=False, hidden=True)
        """
        self.locked = locked
        self.hidden = hidden

class Style:
    """
    Represents formatting settings for a cell or range of cells.
    
    The Style class provides a comprehensive set of properties for controlling the appearance
    of cells, including font, fill, borders, alignment, number format, and protection.
    
    Examples:
        >>> from aspose_cells import Style
        >>> style = Style()
        >>> style.font.bold = True
        >>> style.fill.set_solid_fill('FFFF0000')
        >>> style.borders.set_border('all', 'thin', 'FF000000')
        >>> style.alignment.horizontal = 'center'
    """
    
    def __init__(self):
        """
        Initializes a new instance of Style class with default formatting.
        
        Creates a new Style with default values for all formatting properties.
        
        Examples:
            >>> style = Style()
        """
        self.font = Font()
        self.fill = Fill()
        self.borders = Borders()
        self.alignment = Alignment()
        self.number_format = 'General'
        self.protection = Protection()

    def copy(self):
        """
        Creates a deep copy of this Style object.
        
        Returns:
            Style: A new Style object with the same properties as this Style.
            
        Examples:
            >>> style1 = Style()
            >>> style1.font.bold = True
            >>> style2 = style1.copy()
            >>> style2.font.italic = True  # Doesn't affect style1
        """
        new_style = Style()
        new_style.font = Font(**vars(self.font))
        new_style.fill = Fill(**vars(self.fill))
        new_style.borders = Borders()
        new_style.borders.top = Border(**vars(self.borders.top))
        new_style.borders.bottom = Border(**vars(self.borders.bottom))
        new_style.borders.left = Border(**vars(self.borders.left))
        new_style.borders.right = Border(**vars(self.borders.right))
        new_style.borders.diagonal = Border(**vars(self.borders.diagonal))
        new_style.borders.diagonal_up = self.borders.diagonal_up
        new_style.borders.diagonal_down = self.borders.diagonal_down
        new_style.alignment = Alignment(**vars(self.alignment))
        new_style.number_format = self.number_format
        new_style.protection = Protection(**vars(self.protection))
        return new_style
    
    def set_fill_color(self, color):
        """
        Sets the cell fill color using a solid fill pattern.
        
        Args:
            color (str): Fill color in AARRGGBB hex format.
            
        Examples:
            >>> style.set_fill_color('FFFF0000')  # Red background
        """
        self.fill.set_solid_fill(color)
    
    def set_fill_pattern(self, pattern_type, fg_color='FFFFFFFF', bg_color='FFFFFFFF'):
        """
        Sets the cell fill pattern and colors.
        
        Args:
            pattern_type (str): Pattern type (e.g., 'solid', 'gray125', 'lightGrid').
            fg_color (str, optional): Foreground color in AARRGGBB hex format. Defaults to 'FFFFFFFF'.
            bg_color (str, optional): Background color in AARRGGBB hex format. Defaults to 'FFFFFFFF'.
            
        Examples:
            >>> style.set_fill_pattern('gray125')
            >>> style.set_fill_pattern('lightGrid', 'FFCCCCCC', 'FFFFFFFF')
        """
        self.fill.set_pattern_fill(pattern_type, fg_color, bg_color)
    
    def set_no_fill(self):
        """
        Removes the cell fill (transparent background).
        
        Examples:
            >>> style.set_no_fill()
        """
        self.fill.set_no_fill()
    
    def set_border_color(self, side, color):
        """
        Sets the border color for a specific side.
        
        Args:
            side (str): Border side. Valid values: 'top', 'bottom', 'left', 'right', 'all'.
            color (str): Border color in AARRGGBB hex format.
            
        Examples:
            >>> style.set_border_color('top', 'FFFF0000')  # Red top border
            >>> style.set_border_color('all', 'FF000000')  # Black border on all sides
        """
        if side == 'top':
            self.borders.top.color = color
        elif side == 'bottom':
            self.borders.bottom.color = color
        elif side == 'left':
            self.borders.left.color = color
        elif side == 'right':
            self.borders.right.color = color
        elif side == 'all':
            self.borders.top.color = color
            self.borders.bottom.color = color
            self.borders.left.color = color
            self.borders.right.color = color
    
    def set_border_style(self, side, style):
        """
        Sets the border line style for a specific side.
        
        Args:
            side (str): Border side. Valid values: 'top', 'bottom', 'left', 'right', 'all'.
            style (str): Border line style. Valid values: 'none', 'thin', 'medium', 'thick',
                        'dotted', 'dashed', 'double', 'hair', 'mediumDashed', 'dashDot',
                        'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'.
            
        Examples:
            >>> style.set_border_style('top', 'thin')
            >>> style.set_border_style('all', 'medium')
        """
        if side == 'top':
            self.borders.top.line_style = style
        elif side == 'bottom':
            self.borders.bottom.line_style = style
        elif side == 'left':
            self.borders.left.line_style = style
        elif side == 'right':
            self.borders.right.line_style = style
        elif side == 'all':
            self.borders.top.line_style = style
            self.borders.bottom.line_style = style
            self.borders.left.line_style = style
            self.borders.right.line_style = style
    
    def set_border_weight(self, side, weight):
        """
        Sets the border line weight for a specific side.
        
        Args:
            side (str): Border side. Valid values: 'top', 'bottom', 'left', 'right', 'all'.
            weight (int): Border weight/thickness.
            
        Examples:
            >>> style.set_border_weight('top', 2)
            >>> style.set_border_weight('all', 1)
        """
        if side == 'top':
            self.borders.top.weight = weight
        elif side == 'bottom':
            self.borders.bottom.weight = weight
        elif side == 'left':
            self.borders.left.weight = weight
        elif side == 'right':
            self.borders.right.weight = weight
        elif side == 'all':
            self.borders.top.weight = weight
            self.borders.bottom.weight = weight
            self.borders.left.weight = weight
            self.borders.right.weight = weight
    
    def set_border(self, side, line_style='none', color='FF000000', weight=1):
        """
        Sets complete border properties for a specific side.
        
        Args:
            side (str): Border side. Valid values: 'top', 'bottom', 'left', 'right', 'all'.
            line_style (str, optional): Border line style. Defaults to 'none'.
            color (str, optional): Border color in AARRGGBB hex format. Defaults to 'FF000000'.
            weight (int, optional): Border weight/thickness. Defaults to 1.
            
        Examples:
            >>> style.set_border('top', 'thin', 'FFFF0000', 1)  # Thin red top border
            >>> style.set_border('all', 'medium', 'FF000000', 2)  # Medium black border on all sides
        """
        if side == 'top':
            self.borders.top.line_style = line_style
            self.borders.top.color = color
            self.borders.top.weight = weight
        elif side == 'bottom':
            self.borders.bottom.line_style = line_style
            self.borders.bottom.color = color
            self.borders.bottom.weight = weight
        elif side == 'left':
            self.borders.left.line_style = line_style
            self.borders.left.color = color
            self.borders.right.weight = weight
        elif side == 'right':
            self.borders.right.line_style = line_style
            self.borders.right.color = color
            self.borders.right.weight = weight
        elif side == 'all':
            self.set_border('top', line_style, color, weight)
            self.set_border('bottom', line_style, color, weight)
            self.set_border('left', line_style, color, weight)
            self.set_border('right', line_style, color, weight)
    
    def set_diagonal_border(self, line_style='none', color='FF000000', weight=1, up=False, down=False):
        """
        Sets diagonal border properties.
        
        Args:
            line_style (str, optional): Border line style. Defaults to 'none'.
            color (str, optional): Border color in AARRGGBB hex format..
            weight (int, optional): Border weight/thickness. Defaults to 1.
            up (bool, optional): Whether diagonal border goes from bottom-left to top-right. Defaults to False.
            down (bool, optional): Whether diagonal border goes from top-left to bottom-right. Defaults to False.
            
        Examples:
            >>> style.set_diagonal_border('thin', 'FF000000', 1, up=True, down=True)
        """
        self.borders.diagonal.line_style = line_style
        self.borders.diagonal.color = color
        self.borders.diagonal.weight = weight
        self.borders.diagonal_up = up
        self.borders.diagonal_down = down
    
    def set_horizontal_alignment(self, alignment):
        """
        Sets the horizontal alignment of cell content.
        
        Args:
            alignment (str): Horizontal alignment. Valid values: 'general', 'left', 'center', 'right',
                           'fill', 'justify', 'centerContinuous', 'distributed'.
                           
        Raises:
            ValueError: If alignment is not a valid value.
            
        Examples:
            >>> style.set_horizontal_alignment('center')
            >>> style.set_horizontal_alignment('right')
        """
        valid_alignments = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
        if alignment in valid_alignments:
            self.alignment.horizontal = alignment
        else:
            raise ValueError(f"Invalid horizontal alignment: {alignment}. Valid values: {valid_alignments}")
    
    def set_vertical_alignment(self, alignment):
        """
        Sets the vertical alignment of cell content.
        
        Args:
            alignment (str): Vertical alignment. Valid values: 'top', 'center', 'bottom', 'justify', 'distributed'.
            
        Raises:
            ValueError: If alignment is not a valid value.
            
        Examples:
            >>> style.set_vertical_alignment('center')
            >>> style.set_vertical_alignment('top')
        """
        valid_alignments = ['top', 'center', 'bottom', 'justify', 'distributed']
        if alignment in valid_alignments:
            self.alignment.vertical = alignment
        else:
            raise ValueError(f"Invalid vertical alignment: {alignment}. Valid values: {valid_alignments}")
    
    def set_text_wrap(self, wrap=True):
        """
        Sets whether text wraps within the cell.
        
        Args:
            wrap (bool): True to enable text wrapping, False to disable. Defaults to True.
            
        Examples:
            >>> style.set_text_wrap(True)
            >>> style.set_text_wrap(False)
        """
        self.alignment.wrap_text = wrap
    
    def set_shrink_to_fit(self, shrink=True):
        """
        Sets whether text shrinks to fit the cell width.
        
        Args:
            shrink (bool): True to enable shrink to fit, False to disable. Defaults to True.
            
        Examples:
            >>> style.set_shrink_to_fit(True)
        """
        self.alignment.shrink_to_fit = shrink
    
    def set_indent(self, indent):
        """
        Sets the indent level for cell content.
        
        Args:
            indent (int): Indent level (0-250). Negative values are set to 0.
            
        Examples:
            >>> style.set_indent(2)
            >>> style.set_indent(5)
        """
        self.alignment.indent = max(0, indent)
    
    def set_text_rotation(self, rotation):
        """
        Sets the text rotation angle.
        
        Args:
            rotation (int): Rotation angle in degrees. Valid values: 0-180, or 255 for vertical text.
                          0-90: Rotate text counterclockwise
                          90-180: Rotate text clockwise
                          255: Vertical text
                          
        Raises:
            ValueError: If rotation is not a valid value.
            
        Examples:
            >>> style.set_text_rotation(45)  # 45 degrees counterclockwise
            >>> style.set_text_rotation(255)  # Vertical text
        """
        if rotation == 255 or (0 <= rotation <= 180):
            self.alignment.text_rotation = rotation
        else:
            raise ValueError("Text rotation must be 0-180 degrees or 255 for vertical text")
    
    def set_reading_order(self, order):
        """
        Sets the reading order for cell content.
        
        Args:
            order (int): Reading order. Valid values:
                         0: Context (determined by system)
                         1: Left-to-Right
                         2: Right-to-Left
                         
        Raises:
            ValueError: If order is not a valid value.
            
        Examples:
            >>> style.set_reading_order(1)  # Left-to-Right
            >>> style.set_reading_order(2)  # Right-to-Left
        """
        if order in (0, 1, 2):
            self.alignment.reading_order = order
        else:
            raise ValueError("Reading order must be 0 (Context), 1 (Left-to-Right), or 2 (Right-to-Left)")
    
    def set_number_format(self, format_code):
        """
        Sets the number format code for the cell.
        
        Args:
            format_code (str): Number format code (e.g., '0.00', '#,##0.00', '0%', 'mm/dd/yyyy').
            
        Examples:
            >>> style.set_number_format('0.00')  # Two decimal places
            >>> style.set_number_format('#,##0.00')  # Thousands separator
            >>> style.set_number_format('0%')  # Percentage
            >>> style.set_number_format('mm/dd/yyyy')  # Date format
        """
        self.number_format = format_code
    
    def set_builtin_number_format(self, format_id):
        """
        Sets the number format using a built-in format ID.
        
        Args:
            format_id (int): Built-in format ID (0-164). See NumberFormat.BUILTIN_FORMATS for available IDs.
            
        Examples:
            >>> style.set_builtin_number_format(2)  # '0.00'
            >>> style.set_builtin_number_format(4)  # '#,##0.00'
            >>> style.set_builtin_number_format(14)  # 'mm-dd-yy'
        """
        self.number_format = NumberFormat.get_builtin_format(format_id)
    
    def set_locked(self, locked=True):
        """
        Sets whether the cell is locked when the worksheet is protected.
        
        Args:
            locked (bool): True to lock the cell, False to unlock. Defaults to True.
            
        Examples:
            >>> style.set_locked(True)  # Cell is locked when sheet is protected
            >>> style.set_locked(False)  # Cell can be edited even when sheet is protected
        """
        self.protection.locked = locked
    
    def set_formula_hidden(self, hidden=True):
        """
        Sets whether the cell's formula is hidden when the worksheet is protected.
        
        Args:
            hidden (bool): True to hide the formula, False to show it. Defaults to True.
            
        Examples:
            >>> style.set_formula_hidden(True)  # Formula is hidden when sheet is protected
        """
        self.protection.hidden = hidden