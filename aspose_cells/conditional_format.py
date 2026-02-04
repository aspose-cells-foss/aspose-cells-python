"""
Aspose.Cells for Python - Conditional Format Module

This module provides ConditionalFormat and ConditionalFormatCollection classes which represent
conditional formatting rules in Excel worksheets.

Compatible with Aspose.Cells for .NET API structure and ECMA-376 specification.
"""


class ConditionalFormat:
    """
    Represents a single conditional formatting rule applied to a cell range.
    
    The ConditionalFormat class provides properties for defining conditional formatting rules
    including cell value comparisons, text rules, date rules, duplicate/unique values,
    top/bottom rules, above/below average, color scales, data bars, icon sets, and formula rules.
    
    Examples:
        >>> from aspose_cells import Workbook, ConditionalFormat
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> cf = ws.conditional_formats.add()
        >>> cf.type = "cellValue"
        >>> cf.operator = "greaterThan"
        >>> cf.formula1 = 100
        >>> cf.range = "A1:A10"
        >>> cf.font.bold = True
        >>> cf.font.color = "FFFF0000"
    """
    
    def __init__(self):
        """
        Initializes a new instance of ConditionalFormat class.
        
        Creates a new conditional format with default values for all properties.
        Rule-specific properties are initialized to None and should be set
        based on the conditional format type.
        """
        # Common properties
        self._type = None  # Rule type (cellValue, text, date, duplicateValues, top10, aboveAverage, colorScale, dataBar, iconSet, formula)
        self._range = None  # Cell range in A1 notation (e.g., "A1:A10")
        self._stop_if_true = False  # Stops lower priority rules if this rule is true
        self._priority = 1  # Rule priority (lower = higher priority)
        
        # Rule-specific properties
        self._operator = None  # Comparison operator (for cellValue, text, date rules)
        self._formula1 = None  # First value/formula for comparison
        self._formula2 = None  # Second value/formula (for between operators)
        self._duplicate = None  # Boolean for duplicate/unique values
        self._top = None  # Boolean for top/bottom rules
        self._percent = None  # Boolean for percentage (top/bottom rules)
        self._rank = None  # Integer rank value (top/bottom rules)
        self._above = None  # Boolean for above/below average
        self._std_dev = None  # Integer standard deviations (above/below average)
        self._color_scale_type = None  # '2-color' or '3-color'
        self._min_color = None  # RGB hex for minimum value (color scales)
        self._mid_color = None  # RGB hex for midpoint (3-color scales)
        self._max_color = None  # RGB hex for maximum value (color scales)
        self._bar_color = None  # RGB hex for bar color (data bars)
        self._negative_color = None  # RGB hex for negative values (data bars)
        self._show_border = False  # Show border (data bars)
        self._direction = None  # 'left-to-right' or 'right-to-left' (data bars)
        self._bar_length = None  # 'auto' or fixed value (data bars)
        self._icon_set_type = None  # Icon set name (icon sets)
        self._reverse_icons = False  # Reverse icon order (icon sets)
        self._show_icon_only = False  # Show only icons (icon sets)
        self._formula = None  # Excel formula string (formula rules)
        
        # Formatting properties (applied when condition is true)
        from .style import Font, Fill, Border, Borders, Alignment, Protection
        self._font = Font()
        self._border = Border()
        self._fill = Fill()
        self._alignment = Alignment()
        self._number_format = 'General'
    
    # Common properties
    
    @property
    def type(self):
        """
        Gets or sets the conditional format type.
        
        Returns:
            str: Rule type (cellValue, text, date, duplicateValues, top10, aboveAverage, colorScale, dataBar, iconSet, formula).
            
        Examples:
            >>> cf.type = "cellValue"
            >>> print(cf.type)
        """
        # Map internal XML type name to user-friendly type name
        # Excel persists "cellIs" in XML, but users expect "cellValue"
        if self._type == 'cellIs':
            return 'cellValue'
        return self._type
    
    @type.setter
    def type(self, value):
        """
        Sets the conditional format type.
        
        Args:
            value (str): Rule type.
            
        Raises:
            ValueError: If type is not a valid conditional format type.
        """
        valid_types = [
            'cellValue', 'cellIs',  # cellIs is the XML name for cellValue
            'text', 'containsText', 'notContainsText', 'beginsWith', 'endsWith',  # text rule variants
            'date', 'timePeriod',  # timePeriod is the XML name for date rules
            'duplicateValues', 'uniqueValues',
            'top10', 'bottom10', 'aboveAverage', 'belowAverage',
            'colorScale', 'dataBar', 'iconSet', 'formula', 'expression'  # expression is alias for formula
        ]
        if value in valid_types:
            self._type = value
        else:
            raise ValueError(f"Invalid conditional format type: {value}. Valid types: {valid_types}")
    
    @property
    def range(self):
        """
        Gets or sets the cell range for the conditional format.
        
        Returns:
            str or None: Cell range in A1 notation (e.g., "A1:A10").
            
        Examples:
            >>> cf.range = "A1:A10"
            >>> print(cf.range)
        """
        return self._range
    
    @range.setter
    def range(self, value):
        """
        Sets the cell range for the conditional format.
        
        Args:
            value (str): Cell range in A1 notation.
            
        Examples:
            >>> cf.range = "A1:A10"
        """
        self._range = value
    
    @property
    def stop_if_true(self):
        """
        Gets or sets whether to stop evaluating lower priority rules if this rule is true.
        
        Returns:
            bool: True to stop lower priority rules, False otherwise.
            
        Examples:
            >>> cf.stop_if_true = True
            >>> print(cf.stop_if_true)
        """
        return self._stop_if_true
    
    @stop_if_true.setter
    def stop_if_true(self, value):
        """
        Sets whether to stop evaluating lower priority rules if this rule is true.
        
        Args:
            value (bool): True to stop, False to continue.
            
        Examples:
            >>> cf.stop_if_true = True
        """
        self._stop_if_true = value
    
    @property
    def priority(self):
        """
        Gets or sets the rule priority.
        
        Returns:
            int: Rule priority (lower = higher priority).
            
        Examples:
            >>> cf.priority = 1
            >>> print(cf.priority)
        """
        return self._priority
    
    @priority.setter
    def priority(self, value):
        """
        Sets the rule priority.
        
        Args:
            value (int): Rule priority (positive integer).
            
        Examples:
            >>> cf.priority = 1
        """
        if isinstance(value, int) and value > 0:
            self._priority = value
        else:
            raise ValueError("Priority must be a positive integer")
    
    # Cell Value Rule properties
    
    @property
    def operator(self):
        """
        Gets or sets the comparison operator for cell value rules.
        
        Returns:
            str or None: Comparison operator (equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, between, notBetween).
            
        Examples:
            >>> cf.operator = "greaterThan"
            >>> print(cf.operator)
        """
        return self._operator
    
    @operator.setter
    def operator(self, value):
        """
        Sets the comparison operator for cell value rules.
        
        Args:
            value (str): Comparison operator.
            
        Raises:
            ValueError: If operator is not valid.
        """
        valid_operators = [
            'equal', 'notEqual', 'greaterThan', 'lessThan',
            'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'
        ]
        if value in valid_operators:
            self._operator = value
        else:
            raise ValueError(f"Invalid operator: {value}. Valid operators: {valid_operators}")
    
    @property
    def formula1(self):
        """
        Gets or sets the first value/formula for comparison.
        
        Returns:
            str or None: First value/formula.
            
        Examples:
            >>> cf.formula1 = "100"
            >>> cf.formula1 = "=A1"
        """
        return self._formula1
    
    @formula1.setter
    def formula1(self, value):
        """
        Sets the first value/formula for comparison.
        
        Args:
            value (str): Value or formula string.
            
        Examples:
            >>> cf.formula1 = "100"
            >>> cf.formula1 = "=A1"
        """
        self._formula1 = value
    
    @property
    def formula2(self):
        """
        Gets or sets the second value/formula for comparison.
        
        Returns:
            str or None: Second value/formula (for 'between' and 'notBetween' operators).
            
        Examples:
            >>> cf.formula2 = "200"
            >>> cf.formula2 = "=B1"
        """
        return self._formula2
    
    @formula2.setter
    def formula2(self, value):
        """
        Sets the second value/formula for comparison.
        
        Args:
            value (str): Value or formula string.
            
        Examples:
            >>> cf.formula2 = "200"
            >>> cf.formula2 = "=B1"
        """
        self._formula2 = value
    
    # Text Rule properties
    
    @property
    def text_operator(self):
        """
        Gets or sets the text operator for text rules.
        
        Returns:
            str or None: Text operator (contains, notContains, beginsWith, endsWith).
            
        Examples:
            >>> cf.text_operator = "contains"
            >>> print(cf.text_operator)
        """
        return self._operator
    
    @text_operator.setter
    def text_operator(self, value):
        """
        Sets the text operator for text rules.
        
        Args:
            value (str): Text operator.
            
        Raises:
            ValueError: If operator is not valid.
        """
        valid_operators = ['contains', 'notContains', 'beginsWith', 'endsWith']
        if value in valid_operators:
            self._operator = value
        else:
            raise ValueError(f"Invalid text operator: {value}. Valid operators: {valid_operators}")
    
    @property
    def text_formula(self):
        """
        Gets or sets the text value for comparison.
        
        Returns:
            str or None: Text value to compare.
            
        Examples:
            >>> cf.text_formula = "error"
            >>> print(cf.text_formula)
        """
        return self._formula1
    
    @text_formula.setter
    def text_formula(self, value):
        """
        Sets the text value for comparison.
        
        Args:
            value (str): Text value to compare.
            
        Examples:
            >>> cf.text_formula = "error"
        """
        self._formula1 = value
    
    # Date Rule properties
    
    @property
    def date_operator(self):
        """
        Gets or sets the date operator for date rules.
        
        Returns:
            str or None: Date operator (yesterday, today, tomorrow, last7Days, lastWeek, thisWeek, nextWeek, lastMonth, thisMonth, nextMonth, lastQuarter, thisQuarter, nextQuarter, lastYear, thisYear, nextYear, yearToDate).
            
        Examples:
            >>> cf.date_operator = "today"
            >>> print(cf.date_operator)
        """
        return self._operator
    
    @date_operator.setter
    def date_operator(self, value):
        """
        Sets the date operator for date rules.
        
        Args:
            value (str): Date operator.
            
        Raises:
            ValueError: If operator is not valid.
        """
        valid_operators = [
            'yesterday', 'today', 'tomorrow', 'last7Days', 'lastWeek', 'thisWeek', 'nextWeek',
            'lastMonth', 'thisMonth', 'nextMonth', 'lastQuarter', 'thisQuarter', 'nextQuarter',
            'lastYear', 'thisYear', 'nextYear', 'yearToDate'
        ]
        if value in valid_operators:
            self._operator = value
        else:
            raise ValueError(f"Invalid date operator: {value}. Valid operators: {valid_operators}")
    
    @property
    def date_formula(self):
        """
        Gets or sets the optional date value for comparison.
        
        Returns:
            str or None: Date value for comparison.
            
        Examples:
            >>> cf.date_formula = "2026-01-15"
            >>> print(cf.date_formula)
        """
        return self._formula1
    
    @date_formula.setter
    def date_formula(self, value):
        """
        Sets the optional date value for comparison.
        
        Args:
            value (str): Date value for comparison.
            
        Examples:
            >>> cf.date_formula = "2026-01-15"
        """
        self._formula1 = value
    
    # Duplicate/Unique Values properties
    
    @property
    def duplicate(self):
        """
        Gets or sets whether to highlight duplicate values.
        
        Returns:
            bool or None: True for duplicate, False for unique.
            
        Examples:
            >>> cf.duplicate = True
            >>> print(cf.duplicate)
        """
        return self._duplicate
    
    @duplicate.setter
    def duplicate(self, value):
        """
        Sets whether to highlight duplicate values.
        
        Args:
            value (bool): True for duplicate, False for unique.
            
        Examples:
            >>> cf.duplicate = True
        """
        self._duplicate = value
    
    # Top/Bottom Rule properties
    
    @property
    def top(self):
        """
        Gets or sets whether to highlight top values.
        
        Returns:
            bool or None: True for top, False for bottom.
            
        Examples:
            >>> cf.top = True
            >>> print(cf.top)
        """
        return self._top
    
    @top.setter
    def top(self, value):
        """
        Sets whether to highlight top values.
        
        Args:
            value (bool): True for top, False for bottom.
            
        Examples:
            >>> cf.top = True
        """
        self._top = value
    
    @property
    def percent(self):
        """
        Gets or sets whether rank is a percentage.
        
        Returns:
            bool or None: True for percentage, False for count.
            
        Examples:
            >>> cf.percent = True
            >>> print(cf.percent)
        """
        return self._percent
    
    @percent.setter
    def percent(self, value):
        """
        Sets whether rank is a percentage.
        
        Args:
            value (bool): True for percentage, False for count.
            
        Examples:
            >>> cf.percent = True
        """
        self._percent = value
    
    @property
    def rank(self):
        """
        Gets or sets the rank value.
        
        Returns:
            int or None: Rank value (e.g., 10 for top 10).
            
        Examples:
            >>> cf.rank = 10
            >>> print(cf.rank)
        """
        return self._rank
    
    @rank.setter
    def rank(self, value):
        """
        Sets the rank value.
        
        Args:
            value (int): Rank value.
            
        Examples:
            >>> cf.rank = 10
        """
        self._rank = value
    
    # Above/Below Average properties
    
    @property
    def above(self):
        """
        Gets or sets whether to highlight above average.
        
        Returns:
            bool or None: True for above average, False for below average.
            
        Examples:
            >>> cf.above = True
            >>> print(cf.above)
        """
        return self._above
    
    @above.setter
    def above(self, value):
        """
        Sets whether to highlight above average.
        
        Args:
            value (bool): True for above average, False for below average.
            
        Examples:
            >>> cf.above = True
        """
        self._above = value
    
    @property
    def std_dev(self):
        """
        Gets or sets the number of standard deviations.
        
        Returns:
            int or None: Number of standard deviations.
            
        Examples:
            >>> cf.std_dev = 1
            >>> print(cf.std_dev)
        """
        return self._std_dev
    
    @std_dev.setter
    def std_dev(self, value):
        """
        Sets the number of standard deviations.
        
        Args:
            value (int): Number of standard deviations.
            
        Examples:
            >>> cf.std_dev = 1
        """
        self._std_dev = value
    
    # Color Scale properties
    
    @property
    def color_scale_type(self):
        """
        Gets or sets the color scale type.
        
        Returns:
            str or None: '2-color' or '3-color'.
            
        Examples:
            >>> cf.color_scale_type = "3-color"
            >>> print(cf.color_scale_type)
        """
        return self._color_scale_type
    
    @color_scale_type.setter
    def color_scale_type(self, value):
        """
        Sets the color scale type.
        
        Args:
            value (str): '2-color' or '3-color'.
            
        Raises:
            ValueError: If type is not valid.
        """
        if value in ('2-color', '3-color'):
            self._color_scale_type = value
        else:
            raise ValueError("Color scale type must be '2-color' or '3-color'")
    
    @property
    def min_color(self):
        """
        Gets or sets the minimum value color for color scales.
        
        Returns:
            str or None: RGB hex color string.
            
        Examples:
            >>> cf.min_color = "FF63C384"  # Red
            >>> print(cf.min_color)
        """
        return self._min_color
    
    @min_color.setter
    def min_color(self, value):
        """
        Sets the minimum value color for color scales.
        
        Args:
            value (str): RGB hex color string in AARRGGBB format.
            
        Examples:
            >>> cf.min_color = "FF63C384"  # Red
        """
        self._min_color = value
    
    @property
    def mid_color(self):
        """
        Gets or sets the midpoint color for 3-color scales.
        
        Returns:
            str or None: RGB hex color string.
            
        Examples:
            >>> cf.mid_color = "FFFFEB84"  # Yellow
            >>> print(cf.mid_color)
        """
        return self._mid_color
    
    @mid_color.setter
    def mid_color(self, value):
        """
        Sets the midpoint color for 3-color scales.
        
        Args:
            value (str): RGB hex color string in AARRGGBB format.
            
        Examples:
            >>> cf.mid_color = "FFFFEB84"  # Yellow
        """
        self._mid_color = value
    
    @property
    def max_color(self):
        """
        Gets or sets the maximum value color for color scales.
        
        Returns:
            str or None: RGB hex color string.
            
        Examples:
            >>> cf.max_color = "FF006100"  # Green
            >>> print(cf.max_color)
        """
        return self._max_color
    
    @max_color.setter
    def max_color(self, value):
        """
        Sets the maximum value color for color scales.
        
        Args:
            value (str): RGB hex color string in AARRGGBB format.
            
        Examples:
            >>> cf.max_color = "FF006100"  # Green
        """
        self._max_color = value
    
    # Data Bar properties
    
    @property
    def bar_color(self):
        """
        Gets or sets the bar color for data bars.
        
        Returns:
            str or None: RGB hex color string.
            
        Examples:
            >>> cf.bar_color = "FF006100"  # Green
            >>> print(cf.bar_color)
        """
        return self._bar_color
    
    @bar_color.setter
    def bar_color(self, value):
        """
        Sets the bar color for data bars.
        
        Args:
            value (str): RGB hex color string in AARRGGBB format.
            
        Examples:
            >>> cf.bar_color = "FF006100"  # Green
        """
        self._bar_color = value
    
    @property
    def negative_color(self):
        """
        Gets or sets the negative value color for data bars.
        
        Returns:
            str or None: RGB hex color string.
            
        Examples:
            >>> cf.negative_color = "FFFF0000"  # Red
            >>> print(cf.negative_color)
        """
        return self._negative_color
    
    @negative_color.setter
    def negative_color(self, value):
        """
        Sets the negative value color for data bars.
        
        Args:
            value (str): RGB hex color string in AARRGGBB format.
            
        Examples:
            >>> cf.negative_color = "FFFF0000"  # Red
        """
        self._negative_color = value
    
    @property
    def show_border(self):
        """
        Gets or sets whether to show border for data bars.
        
        Returns:
            bool: Show border.
            
        Examples:
            >>> cf.show_border = True
            >>> print(cf.show_border)
        """
        return self._show_border
    
    @show_border.setter
    def show_border(self, value):
        """
        Sets whether to show border for data bars.
        
        Args:
            value (bool): True to show, False to hide.
            
        Examples:
            >>> cf.show_border = True
        """
        self._show_border = value
    
    @property
    def direction(self):
        """
        Gets or sets the bar direction for data bars.
        
        Returns:
            str or None: 'left-to-right' or 'right-to-left'.
            
        Examples:
            >>> cf.direction = "left-to-right"
            >>> print(cf.direction)
        """
        return self._direction
    
    @direction.setter
    def direction(self, value):
        """
        Sets the bar direction for data bars.
        
        Args:
            value (str): 'left-to-right' or 'right-to-left'.
            
        Raises:
            ValueError: If direction is not valid.
        """
        if value in ('left-to-right', 'right-to-left'):
            self._direction = value
        else:
            raise ValueError("Direction must be 'left-to-right' or 'right-to-left'")
    
    @property
    def bar_length(self):
        """
        Gets or sets the bar length for data bars.
        
        Returns:
            str or None: 'auto' or fixed length value.
            
        Examples:
            >>> cf.bar_length = "auto"
            >>> print(cf.bar_length)
        """
        return self._bar_length
    
    @bar_length.setter
    def bar_length(self, value):
        """
        Sets the bar length for data bars.
        
        Args:
            value (str): 'auto' or fixed length value.
            
        Examples:
            >>> cf.bar_length = "auto"
            >>> cf.bar_length = 50
        """
        self._bar_length = value
    
    # Icon Set properties
    
    @property
    def icon_set_type(self):
        """
        Gets or sets the icon set type.
        
        Returns:
            str or None: Icon set name.
            
        Examples:
            >>> cf.icon_set_type = "3TrafficLights1"
            >>> print(cf.icon_set_type)
        """
        return self._icon_set_type
    
    @icon_set_type.setter
    def icon_set_type(self, value):
        """
        Sets the icon set type.
        
        Args:
            value (str): Icon set name.
            
        Raises:
            ValueError: If icon set type is not valid.
        """
        valid_sets = [
            '3Arrows', '3TrafficLights1', '3TrafficLights2', '3Flags', '3Signs',
            '4Arrows', '4ArrowsGray', '4TrafficLights', '5Arrows', '5ArrowsGray',
            '5Quarters', '5Rating', '5Symbols', '3Symbols', '3Symbols2'
        ]
        if value in valid_sets:
            self._icon_set_type = value
        else:
            raise ValueError(f"Invalid icon set type: {value}. Valid types: {valid_sets}")
    
    @property
    def reverse_icons(self):
        """
        Gets or sets whether to reverse icon order.
        
        Returns:
            bool: Reverse icon order.
            
        Examples:
            >>> cf.reverse_icons = True
            >>> print(cf.reverse_icons)
        """
        return self._reverse_icons
    
    @reverse_icons.setter
    def reverse_icons(self, value):
        """
        Sets whether to reverse icon order.
        
        Args:
            value (bool): True to reverse, False for normal.
            
        Examples:
            >>> cf.reverse_icons = True
        """
        self._reverse_icons = value
    
    @property
    def show_icon_only(self):
        """
        Gets or sets whether to show only icons.
        
        Returns:
            bool: Show only icons (hide values).
            
        Examples:
            >>> cf.show_icon_only = True
            >>> print(cf.show_icon_only)
        """
        return self._show_icon_only
    
    @show_icon_only.setter
    def show_icon_only(self, value):
        """
        Sets whether to show only icons.
        
        Args:
            value (bool): True to show only icons, False to show values.
            
        Examples:
            >>> cf.show_icon_only = True
        """
        self._show_icon_only = value
    
    # Formula Rule properties
    
    @property
    def formula(self):
        """
        Gets or sets the formula for formula-based rules.
        
        Returns:
            str or None: Excel formula string.
            
        Examples:
            >>> cf.formula = "=A1>100"
            >>> print(cf.formula)
        """
        return self._formula
    
    @formula.setter
    def formula(self, value):
        """
        Sets the formula for formula-based rules.
        
        Args:
            value (str): Excel formula string.
            
        Examples:
            >>> cf.formula = "=A1>100"
        """
        self._formula = value
    
    # Formatting properties (applied when condition is true)
    
    @property
    def font(self):
        """
        Gets the font settings for this conditional format.
        
        Returns:
            Font: Font object.
            
        Examples:
            >>> cf.font.bold = True
            >>> cf.font.color = "FFFF0000"
        """
        return self._font
    
    @property
    def border(self):
        """
        Gets the border settings for this conditional format.
        
        Returns:
            Border: Border object.
            
        Examples:
            >>> cf.border.line_style = "thin"
            >>> cf.border.color = "FFFF0000"
        """
        return self._border
    
    @property
    def fill(self):
        """
        Gets the fill settings for this conditional format.
        
        Returns:
            Fill: Fill object.
            
        Examples:
            >>> cf.fill.set_solid_fill('FFFF0000')
        """
        return self._fill
    
    @property
    def alignment(self):
        """
        Gets the alignment settings for this conditional format.
        
        Returns:
            Alignment: Alignment object.
            
        Examples:
            >>> cf.alignment.horizontal = "center"
        """
        return self._alignment
    
    @property
    def number_format(self):
        """
        Gets or sets the number format for this conditional format.
        
        Returns:
            str: Number format string.
            
        Examples:
            >>> cf.number_format = "0.00"
        """
        return self._number_format
    
    @number_format.setter
    def number_format(self, value):
        """
        Sets the number format for this conditional format.
        
        Args:
            value (str): Number format string.
            
        Examples:
            >>> cf.number_format = "0.00"
        """
        self._number_format = value


class ConditionalFormatCollection:
    """
    Represents a collection of conditional formats for a worksheet.
    
    The ConditionalFormatCollection class provides methods for managing
    multiple conditional formatting rules applied to a worksheet.
    
    Examples:
        >>> from aspose_cells import Workbook
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>> cf = ws.conditional_formats.add()
        >>> cf.type = "cellValue"
        >>> cf.range = "A1:A10"
        >>> print(len(ws.conditional_formats))  # 1
    """
    
    def __init__(self):
        """
        Initializes a new instance of ConditionalFormatCollection class.
        
        Creates an empty collection for conditional formats.
        """
        self._formats = []
    
    @property
    def count(self):
        """
        Gets the number of conditional formats in the collection.
        
        Returns:
            int: Number of conditional formats.
            
        Examples:
            >>> print(len(ws.conditional_formats))
        """
        return len(self._formats)
    
    def add(self):
        """
        Adds a new conditional format to the collection.
        
        Returns:
            ConditionalFormat: The newly created ConditionalFormat object.
            
        Examples:
            >>> cf = ws.conditional_formats.add()
            >>> cf.type = "cellValue"
        """
        cf = ConditionalFormat()
        self._formats.append(cf)
        return cf
    
    def get_by_index(self, index):
        """
        Gets a conditional format by its index.
        
        Args:
            index (int): Zero-based index of the conditional format.
            
        Returns:
            ConditionalFormat or None: The ConditionalFormat object at the specified index, or None if not found.
            
        Examples:
            >>> cf = ws.conditional_formats.get_by_index(0)
            >>> if cf:
            ...     print(cf.type)
        """
        if 0 <= index < len(self._formats):
            return self._formats[index]
        return None
    
    def get_by_range(self, range_str):
        """
        Gets a conditional format by its range.
        
        Args:
            range_str (str): Cell range in A1 notation.
            
        Returns:
            ConditionalFormat or None: The ConditionalFormat object at the specified range, or None if not found.
            
        Examples:
            >>> cf = ws.conditional_formats.get_by_range("A1:A10")
            >>> if cf:
            ...     print(cf.type)
        """
        for cf in self._formats:
            if cf.range == range_str:
                return cf
        return None
    
    def remove(self, cf):
        """
        Removes a conditional format from the collection.
        
        Args:
            cf (ConditionalFormat or int): The ConditionalFormat object or index to remove.
            
        Examples:
            >>> ws.conditional_formats.remove(0)  # Remove by index
            >>> ws.conditional_formats.remove(cf)  # Remove by object
        """
        if isinstance(cf, int):
            if 0 <= cf < len(self._formats):
                del self._formats[cf]
        elif cf in self._formats:
            self._formats.remove(cf)
    
    def clear(self):
        """
        Clears all conditional formats from the collection.
        
        Examples:
            >>> ws.conditional_formats.clear()
        """
        self._formats = []
    
    def __iter__(self):
        """
        Iterates over all conditional formats in the collection.
        
        Yields:
            ConditionalFormat: Each conditional format in the collection.
            
        Examples:
            >>> for cf in ws.conditional_formats:
            ...     print(cf.type)
        """
        return iter(self._formats)
    
    def __len__(self):
        """
        Gets the number of conditional formats in the collection.
        
        Returns:
            int: Number of conditional formats.
        """
        return len(self._formats)
