"""
Aspose.Cells for Python - Data Validation Module

This module provides classes for Excel data validation according to ECMA-376 specification.
Data validation allows you to control what data can be entered into cells.

Compatible with Excel VBA Validation object API design and ECMA-376 SpreadsheetML.

References:
- ECMA-376 Part 4, Section 3.3.1.30 (dataValidation)
- Excel VBA Validation object: https://learn.microsoft.com/en-us/office/vba/api/excel.validation
"""

from enum import IntEnum


class DataValidationType(IntEnum):
    """
    Specifies the type of data validation.

    Corresponds to XlDVType enumeration in Excel VBA and
    ST_DataValidationType in ECMA-376.
    """
    NONE = 0           # No validation (xlValidateInputOnly equivalent)
    WHOLE_NUMBER = 1   # Whole number validation (xlValidateWholeNumber)
    DECIMAL = 2        # Decimal number validation (xlValidateDecimal)
    LIST = 3           # List selection validation (xlValidateList)
    DATE = 4           # Date validation (xlValidateDate)
    TIME = 5           # Time validation (xlValidateTime)
    TEXT_LENGTH = 6    # Text length validation (xlValidateTextLength)
    CUSTOM = 7         # Custom formula validation (xlValidateCustom)


class DataValidationOperator(IntEnum):
    """
    Specifies the comparison operator for data validation.

    Corresponds to XlFormatConditionOperator enumeration in Excel VBA and
    ST_DataValidationOperator in ECMA-376.
    """
    BETWEEN = 0              # Value must be between formula1 and formula2
    NOT_BETWEEN = 1          # Value must NOT be between formula1 and formula2
    EQUAL = 2                # Value must equal formula1
    NOT_EQUAL = 3            # Value must NOT equal formula1
    GREATER_THAN = 4         # Value must be greater than formula1
    LESS_THAN = 5            # Value must be less than formula1
    GREATER_THAN_OR_EQUAL = 6  # Value must be >= formula1
    LESS_THAN_OR_EQUAL = 7     # Value must be <= formula1


class DataValidationAlertStyle(IntEnum):
    """
    Specifies the style of the error alert displayed when invalid data is entered.

    Corresponds to XlDVAlertStyle enumeration in Excel VBA and
    ST_DataValidationErrorStyle in ECMA-376.
    """
    STOP = 0         # Prevents invalid data entry (most restrictive)
    WARNING = 1      # Warns user but allows invalid data if confirmed
    INFORMATION = 2  # Informs user but allows invalid data if confirmed


class DataValidationImeMode(IntEnum):
    """
    Specifies the Input Method Editor (IME) mode for CJK language input.

    Corresponds to ST_DataValidationImeMode in ECMA-376.
    Only applies for Chinese, Japanese, and Korean languages.
    """
    NO_CONTROL = 0      # No IME control (default)
    OFF = 1             # IME off
    ON = 2              # IME on
    DISABLED = 3        # IME disabled
    HIRAGANA = 4        # Japanese Hiragana mode
    FULL_KATAKANA = 5   # Japanese full-width Katakana mode
    HALF_KATAKANA = 6   # Japanese half-width Katakana mode
    FULL_ALPHA = 7      # Full-width alphanumeric mode
    HALF_ALPHA = 8      # Half-width alphanumeric mode
    FULL_HANGUL = 9     # Korean full-width Hangul mode
    HALF_HANGUL = 10    # Korean half-width Hangul mode


class DataValidation:
    """
    Represents data validation settings for a range of cells.

    This class provides properties and methods for defining validation rules
    that control what data can be entered into cells, similar to Excel's
    Validation object.

    Examples:
        >>> from aspose_cells import Workbook
        >>> from aspose_cells.data_validation import DataValidation, DataValidationType
        >>>
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>>
        >>> # Add a whole number validation
        >>> validation = ws.data_validations.add("A1:A10")
        >>> validation.type = DataValidationType.WHOLE_NUMBER
        >>> validation.operator = DataValidationOperator.BETWEEN
        >>> validation.formula1 = "1"
        >>> validation.formula2 = "100"
        >>> validation.error_message = "Please enter a number between 1 and 100"
    """

    def __init__(self, sqref=None):
        """
        Initializes a new DataValidation instance.

        Args:
            sqref (str, optional): Cell range(s) for validation in A1 notation.
                                   Can be a single range like "A1:A10" or multiple
                                   ranges separated by spaces like "A1:A10 C1:C10".
        """
        # Cell range for validation (required attribute in ECMA-376)
        self._sqref = sqref

        # Validation type and operator
        self._type = DataValidationType.NONE
        self._operator = DataValidationOperator.BETWEEN

        # Formulas for validation
        self._formula1 = None
        self._formula2 = None

        # Error alert settings
        self._alert_style = DataValidationAlertStyle.STOP
        self._show_error_message = True
        self._error_title = None
        self._error_message = None

        # Input message settings
        self._show_input_message = False
        self._input_title = None
        self._input_message = None

        # Other settings
        self._allow_blank = True
        self._show_dropdown = True  # Note: In ECMA-376, false = show dropdown (counterintuitive)
        self._ime_mode = DataValidationImeMode.NO_CONTROL

    # ==================== Cell Range ====================

    @property
    def sqref(self):
        """
        Gets or sets the cell range(s) for this validation.

        The range can be a single range (e.g., "A1:A10") or multiple ranges
        separated by spaces (e.g., "A1:A10 C1:C10 E1:E10").

        Returns:
            str: Cell range(s) in A1 notation.
        """
        return self._sqref

    @sqref.setter
    def sqref(self, value):
        """Sets the cell range(s) for this validation."""
        self._sqref = value

    @property
    def ranges(self):
        """
        Gets the cell range(s) as a list.

        Returns:
            list: List of range strings.
        """
        if self._sqref:
            return self._sqref.split()
        return []

    # ==================== Validation Type and Operator ====================

    @property
    def type(self):
        """
        Gets or sets the validation type.

        Returns:
            DataValidationType: The type of validation.

        Examples:
            >>> validation.type = DataValidationType.WHOLE_NUMBER
            >>> validation.type = DataValidationType.LIST
        """
        return self._type

    @type.setter
    def type(self, value):
        """Sets the validation type."""
        if isinstance(value, int):
            self._type = DataValidationType(value)
        else:
            self._type = value

    @property
    def operator(self):
        """
        Gets or sets the comparison operator.

        Note: The operator is only applicable when type is WHOLE_NUMBER, DECIMAL,
        DATE, TIME, or TEXT_LENGTH. For LIST, CUSTOM, and NONE types, the
        operator is ignored.

        Returns:
            DataValidationOperator: The comparison operator.
        """
        return self._operator

    @operator.setter
    def operator(self, value):
        """Sets the comparison operator."""
        if isinstance(value, int):
            self._operator = DataValidationOperator(value)
        else:
            self._operator = value

    # ==================== Formulas ====================

    @property
    def formula1(self):
        """
        Gets or sets the first formula for validation.

        For numeric types: The value or cell reference for comparison.
        For LIST type: Comma-separated values or range reference (e.g., "$A$1:$A$10").
        For CUSTOM type: A formula that returns TRUE/FALSE.

        Returns:
            str: The first formula.

        Examples:
            >>> validation.formula1 = "10"  # Numeric value
            >>> validation.formula1 = "$A$1"  # Cell reference
            >>> validation.formula1 = '"Red,Green,Blue"'  # Inline list
            >>> validation.formula1 = "$A$1:$A$10"  # List from range
            >>> validation.formula1 = "=AND(A1>0, A1<100)"  # Custom formula
        """
        return self._formula1

    @formula1.setter
    def formula1(self, value):
        """Sets the first formula."""
        self._formula1 = str(value) if value is not None else None

    @property
    def formula2(self):
        """
        Gets or sets the second formula for validation.

        Only used when operator is BETWEEN or NOT_BETWEEN.
        Represents the upper bound of the valid range (formula1 is lower bound).

        Returns:
            str: The second formula.
        """
        return self._formula2

    @formula2.setter
    def formula2(self, value):
        """Sets the second formula."""
        self._formula2 = str(value) if value is not None else None

    # ==================== Error Alert Settings ====================

    @property
    def alert_style(self):
        """
        Gets or sets the style of the error alert.

        - STOP: Prevents invalid data entry (default)
        - WARNING: Warns but allows invalid data if confirmed
        - INFORMATION: Informs but allows invalid data

        Returns:
            DataValidationAlertStyle: The alert style.
        """
        return self._alert_style

    @alert_style.setter
    def alert_style(self, value):
        """Sets the alert style."""
        if isinstance(value, int):
            self._alert_style = DataValidationAlertStyle(value)
        else:
            self._alert_style = value

    @property
    def show_error_message(self):
        """
        Gets or sets whether to show the error alert when invalid data is entered.

        Returns:
            bool: True to show error alert, False to hide.
        """
        return self._show_error_message

    @show_error_message.setter
    def show_error_message(self, value):
        """Sets whether to show error alert."""
        self._show_error_message = bool(value)

    # Alias for Excel VBA compatibility
    @property
    def show_error(self):
        """Alias for show_error_message (Excel VBA compatibility)."""
        return self._show_error_message

    @show_error.setter
    def show_error(self, value):
        """Sets show_error (alias for show_error_message)."""
        self._show_error_message = bool(value)

    @property
    def error_title(self):
        """
        Gets or sets the title of the error alert dialog.

        Maximum length is 32 characters.

        Returns:
            str: The error title.
        """
        return self._error_title

    @error_title.setter
    def error_title(self, value):
        """Sets the error title."""
        if value and len(str(value)) > 32:
            value = str(value)[:32]
        self._error_title = str(value) if value else None

    @property
    def error_message(self):
        """
        Gets or sets the error message displayed when invalid data is entered.

        Maximum length is 225 characters.

        Returns:
            str: The error message.
        """
        return self._error_message

    @error_message.setter
    def error_message(self, value):
        """Sets the error message."""
        if value and len(str(value)) > 225:
            value = str(value)[:225]
        self._error_message = str(value) if value else None

    # Alias for ECMA-376 XML attribute name
    @property
    def error(self):
        """Alias for error_message (ECMA-376 compatibility)."""
        return self._error_message

    @error.setter
    def error(self, value):
        """Sets error (alias for error_message)."""
        self.error_message = value

    # ==================== Input Message Settings ====================

    @property
    def show_input_message(self):
        """
        Gets or sets whether to show the input message when the cell is selected.

        Returns:
            bool: True to show input message, False to hide.
        """
        return self._show_input_message

    @show_input_message.setter
    def show_input_message(self, value):
        """Sets whether to show input message."""
        self._show_input_message = bool(value)

    # Alias for Excel VBA compatibility
    @property
    def show_input(self):
        """Alias for show_input_message (Excel VBA compatibility)."""
        return self._show_input_message

    @show_input.setter
    def show_input(self, value):
        """Sets show_input (alias for show_input_message)."""
        self._show_input_message = bool(value)

    @property
    def input_title(self):
        """
        Gets or sets the title of the input message dialog.

        Maximum length is 32 characters.

        Returns:
            str: The input title.
        """
        return self._input_title

    @input_title.setter
    def input_title(self, value):
        """Sets the input title."""
        if value and len(str(value)) > 32:
            value = str(value)[:32]
        self._input_title = str(value) if value else None

    # Alias for ECMA-376 XML attribute name
    @property
    def prompt_title(self):
        """Alias for input_title (ECMA-376 compatibility)."""
        return self._input_title

    @prompt_title.setter
    def prompt_title(self, value):
        """Sets prompt_title (alias for input_title)."""
        self.input_title = value

    @property
    def input_message(self):
        """
        Gets or sets the input message displayed when the cell is selected.

        Maximum length is 255 characters.

        Returns:
            str: The input message.
        """
        return self._input_message

    @input_message.setter
    def input_message(self, value):
        """Sets the input message."""
        if value and len(str(value)) > 255:
            value = str(value)[:255]
        self._input_message = str(value) if value else None

    # Alias for ECMA-376 XML attribute name
    @property
    def prompt(self):
        """Alias for input_message (ECMA-376 compatibility)."""
        return self._input_message

    @prompt.setter
    def prompt(self, value):
        """Sets prompt (alias for input_message)."""
        self.input_message = value

    # ==================== Other Settings ====================

    @property
    def allow_blank(self):
        """
        Gets or sets whether blank/empty entries are valid.

        Corresponds to IgnoreBlank property in Excel VBA.

        Returns:
            bool: True to allow blank entries, False to require data.
        """
        return self._allow_blank

    @allow_blank.setter
    def allow_blank(self, value):
        """Sets whether blank entries are allowed."""
        self._allow_blank = bool(value)

    # Alias for Excel VBA compatibility
    @property
    def ignore_blank(self):
        """Alias for allow_blank (Excel VBA compatibility)."""
        return self._allow_blank

    @ignore_blank.setter
    def ignore_blank(self, value):
        """Sets ignore_blank (alias for allow_blank)."""
        self._allow_blank = bool(value)

    @property
    def show_dropdown(self):
        """
        Gets or sets whether to show the dropdown arrow for list validation.

        Only applicable when type is LIST.

        Note: In ECMA-376 XML, the attribute is named 'showDropDown' but with
        inverted logic (false = show dropdown). This property uses intuitive
        logic (True = show dropdown).

        Returns:
            bool: True to show dropdown, False to hide.
        """
        return self._show_dropdown

    @show_dropdown.setter
    def show_dropdown(self, value):
        """Sets whether to show dropdown."""
        self._show_dropdown = bool(value)

    # Alias for Excel VBA compatibility
    @property
    def in_cell_dropdown(self):
        """Alias for show_dropdown (Excel VBA compatibility)."""
        return self._show_dropdown

    @in_cell_dropdown.setter
    def in_cell_dropdown(self, value):
        """Sets in_cell_dropdown (alias for show_dropdown)."""
        self._show_dropdown = bool(value)

    @property
    def ime_mode(self):
        """
        Gets or sets the IME (Input Method Editor) mode.

        Only applies for Chinese, Japanese, and Korean languages.

        Returns:
            DataValidationImeMode: The IME mode.
        """
        return self._ime_mode

    @ime_mode.setter
    def ime_mode(self, value):
        """Sets the IME mode."""
        if isinstance(value, int):
            self._ime_mode = DataValidationImeMode(value)
        else:
            self._ime_mode = value

    # ==================== Methods ====================

    def add(self, validation_type, alert_style=None, operator=None, formula1=None, formula2=None):
        """
        Configures the data validation with the specified parameters.

        This method mirrors Excel VBA's Validation.Add method signature.

        Args:
            validation_type (DataValidationType or int): The validation type.
            alert_style (DataValidationAlertStyle, optional): The alert style.
            operator (DataValidationOperator, optional): The comparison operator.
            formula1 (str, optional): The first formula.
            formula2 (str, optional): The second formula (for between/notBetween).

        Examples:
            >>> validation.add(DataValidationType.WHOLE_NUMBER,
            ...                DataValidationAlertStyle.STOP,
            ...                DataValidationOperator.BETWEEN, "1", "100")
        """
        self.type = validation_type

        if alert_style is not None:
            self.alert_style = alert_style

        if operator is not None:
            self.operator = operator

        if formula1 is not None:
            self.formula1 = formula1

        if formula2 is not None:
            self.formula2 = formula2

    def modify(self, validation_type=None, alert_style=None, operator=None,
               formula1=None, formula2=None):
        """
        Modifies the data validation settings.

        This method mirrors Excel VBA's Validation.Modify method signature.
        Only provided parameters will be changed.

        Args:
            validation_type (DataValidationType, optional): The validation type.
            alert_style (DataValidationAlertStyle, optional): The alert style.
            operator (DataValidationOperator, optional): The comparison operator.
            formula1 (str, optional): The first formula.
            formula2 (str, optional): The second formula.
        """
        if validation_type is not None:
            self.type = validation_type

        if alert_style is not None:
            self.alert_style = alert_style

        if operator is not None:
            self.operator = operator

        if formula1 is not None:
            self.formula1 = formula1

        if formula2 is not None:
            self.formula2 = formula2

    def delete(self):
        """
        Clears the validation settings (resets to no validation).

        This method mirrors Excel VBA's Validation.Delete method.
        """
        self._type = DataValidationType.NONE
        self._operator = DataValidationOperator.BETWEEN
        self._formula1 = None
        self._formula2 = None
        self._alert_style = DataValidationAlertStyle.STOP
        self._show_error_message = False
        self._error_title = None
        self._error_message = None
        self._show_input_message = False
        self._input_title = None
        self._input_message = None
        self._allow_blank = True
        self._show_dropdown = True
        self._ime_mode = DataValidationImeMode.NO_CONTROL

    def copy(self):
        """
        Creates a copy of this DataValidation.

        Returns:
            DataValidation: A new DataValidation with the same settings.
        """
        new_validation = DataValidation(self._sqref)
        new_validation._type = self._type
        new_validation._operator = self._operator
        new_validation._formula1 = self._formula1
        new_validation._formula2 = self._formula2
        new_validation._alert_style = self._alert_style
        new_validation._show_error_message = self._show_error_message
        new_validation._error_title = self._error_title
        new_validation._error_message = self._error_message
        new_validation._show_input_message = self._show_input_message
        new_validation._input_title = self._input_title
        new_validation._input_message = self._input_message
        new_validation._allow_blank = self._allow_blank
        new_validation._show_dropdown = self._show_dropdown
        new_validation._ime_mode = self._ime_mode
        return new_validation

    def __repr__(self):
        """Returns a string representation of the validation."""
        return (f"DataValidation(sqref='{self._sqref}', type={self._type.name}, "
                f"operator={self._operator.name})")


class DataValidationCollection:
    """
    Represents a collection of DataValidation objects for a worksheet.

    This class manages all data validation rules applied to a worksheet
    and provides methods to add, remove, and access validations.

    Examples:
        >>> from aspose_cells import Workbook
        >>> wb = Workbook()
        >>> ws = wb.worksheets[0]
        >>>
        >>> # Add validations
        >>> dv1 = ws.data_validations.add("A1:A10")
        >>> dv2 = ws.data_validations.add("B1:B10")
        >>>
        >>> # Iterate over validations
        >>> for validation in ws.data_validations:
        ...     print(validation.sqref)
        >>>
        >>> # Access by index
        >>> first = ws.data_validations[0]
        >>>
        >>> # Get count
        >>> count = len(ws.data_validations)
    """

    def __init__(self):
        """Initializes a new DataValidationCollection."""
        self._validations = []
        self._disable_prompts = False
        self._x_window = None
        self._y_window = None

    @property
    def count(self):
        """
        Gets the number of validations in the collection.

        Returns:
            int: The number of validations.
        """
        return len(self._validations)

    @property
    def disable_prompts(self):
        """
        Gets or sets whether all input prompts are disabled.

        Returns:
            bool: True to disable all prompts, False to enable.
        """
        return self._disable_prompts

    @disable_prompts.setter
    def disable_prompts(self, value):
        """Sets whether all prompts are disabled."""
        self._disable_prompts = bool(value)

    @property
    def x_window(self):
        """
        Gets or sets the X coordinate of the dropdown window.

        Returns:
            int or None: X coordinate.
        """
        return self._x_window

    @x_window.setter
    def x_window(self, value):
        """Sets the X coordinate of the dropdown window."""
        self._x_window = int(value) if value is not None else None

    @property
    def y_window(self):
        """
        Gets or sets the Y coordinate of the dropdown window.

        Returns:
            int or None: Y coordinate.
        """
        return self._y_window

    @y_window.setter
    def y_window(self, value):
        """Sets the Y coordinate of the dropdown window."""
        self._y_window = int(value) if value is not None else None

    def add(self, sqref, validation_type=None, operator=None, formula1=None, formula2=None):
        """
        Adds a new data validation to the collection.

        Args:
            sqref (str): Cell range(s) for validation in A1 notation.
            validation_type (DataValidationType, optional): The validation type.
            operator (DataValidationOperator, optional): The comparison operator.
            formula1 (str, optional): The first formula.
            formula2 (str, optional): The second formula.

        Returns:
            DataValidation: The newly created DataValidation object.

        Examples:
            >>> # Add empty validation (configure later)
            >>> dv = ws.data_validations.add("A1:A10")
            >>> dv.type = DataValidationType.WHOLE_NUMBER
            >>>
            >>> # Add with parameters
            >>> dv = ws.data_validations.add("B1:B10",
            ...                              DataValidationType.LIST,
            ...                              formula1='"Red,Green,Blue"')
        """
        validation = DataValidation(sqref)

        if validation_type is not None:
            validation.type = validation_type

        if operator is not None:
            validation.operator = operator

        if formula1 is not None:
            validation.formula1 = formula1

        if formula2 is not None:
            validation.formula2 = formula2

        self._validations.append(validation)
        return validation

    def add_validation(self, validation):
        """
        Adds an existing DataValidation object to the collection.

        Args:
            validation (DataValidation): The validation to add.

        Returns:
            DataValidation: The added validation.
        """
        self._validations.append(validation)
        return validation

    def remove(self, validation):
        """
        Removes a validation from the collection.

        Args:
            validation (DataValidation): The validation to remove.

        Returns:
            bool: True if removed, False if not found.
        """
        if validation in self._validations:
            self._validations.remove(validation)
            return True
        return False

    def remove_at(self, index):
        """
        Removes a validation at the specified index.

        Args:
            index (int): The zero-based index of the validation to remove.

        Raises:
            IndexError: If index is out of range.
        """
        del self._validations[index]

    def remove_by_range(self, sqref):
        """
        Removes validations that match the specified range.

        Args:
            sqref (str): The range to match.

        Returns:
            int: Number of validations removed.
        """
        removed = 0
        self._validations = [v for v in self._validations if v.sqref != sqref]
        return removed

    def clear(self):
        """
        Removes all validations from the collection.
        """
        self._validations = []

    def get_validation(self, cell_ref):
        """
        Gets the validation that applies to a specific cell.

        Args:
            cell_ref (str): Cell reference in A1 notation (e.g., "A1").

        Returns:
            DataValidation or None: The validation if found, None otherwise.
        """
        for validation in self._validations:
            if validation.sqref and self._cell_in_range(cell_ref, validation.sqref):
                return validation
        return None

    def _cell_in_range(self, cell_ref, sqref):
        """
        Checks if a cell is within a range or ranges.

        Args:
            cell_ref (str): Cell reference (e.g., "A1").
            sqref (str): Range(s) (e.g., "A1:A10" or "A1:A10 C1:C10").

        Returns:
            bool: True if cell is in range, False otherwise.
        """
        from .cells import Cells

        # Parse cell reference
        cell_col, cell_row = Cells.parse_cell_reference(cell_ref)

        # Check each range in sqref
        for range_str in sqref.split():
            if ':' in range_str:
                start, end = range_str.split(':')
                start_col, start_row = Cells.parse_cell_reference(start)
                end_col, end_row = Cells.parse_cell_reference(end)

                if (start_col <= cell_col <= end_col and
                    start_row <= cell_row <= end_row):
                    return True
            else:
                ref_col, ref_row = Cells.parse_cell_reference(range_str)
                if cell_col == ref_col and cell_row == ref_row:
                    return True

        return False

    def __len__(self):
        """Returns the number of validations."""
        return len(self._validations)

    def __getitem__(self, index):
        """
        Gets a validation by index.

        Args:
            index (int): Zero-based index.

        Returns:
            DataValidation: The validation at the specified index.
        """
        return self._validations[index]

    def __iter__(self):
        """Returns an iterator over the validations."""
        return iter(self._validations)

    def __repr__(self):
        """Returns a string representation of the collection."""
        return f"DataValidationCollection(count={len(self._validations)})"
