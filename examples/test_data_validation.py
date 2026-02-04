"""
Tests for Data Validation Feature

This module tests the Excel data validation implementation according to ECMA-376 specification.
"""

import sys
import os
import unittest

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import (
    Workbook, DataValidation, DataValidationCollection,
    DataValidationType, DataValidationOperator,
    DataValidationAlertStyle, DataValidationImeMode
)


class TestDataValidationEnums(unittest.TestCase):
    """Tests for data validation enumeration types."""

    def test_data_validation_type_values(self):
        """Test DataValidationType enum values match ECMA-376 spec."""
        self.assertEqual(DataValidationType.NONE, 0)
        self.assertEqual(DataValidationType.WHOLE_NUMBER, 1)
        self.assertEqual(DataValidationType.DECIMAL, 2)
        self.assertEqual(DataValidationType.LIST, 3)
        self.assertEqual(DataValidationType.DATE, 4)
        self.assertEqual(DataValidationType.TIME, 5)
        self.assertEqual(DataValidationType.TEXT_LENGTH, 6)
        self.assertEqual(DataValidationType.CUSTOM, 7)

    def test_data_validation_operator_values(self):
        """Test DataValidationOperator enum values."""
        self.assertEqual(DataValidationOperator.BETWEEN, 0)
        self.assertEqual(DataValidationOperator.NOT_BETWEEN, 1)
        self.assertEqual(DataValidationOperator.EQUAL, 2)
        self.assertEqual(DataValidationOperator.NOT_EQUAL, 3)
        self.assertEqual(DataValidationOperator.GREATER_THAN, 4)
        self.assertEqual(DataValidationOperator.LESS_THAN, 5)
        self.assertEqual(DataValidationOperator.GREATER_THAN_OR_EQUAL, 6)
        self.assertEqual(DataValidationOperator.LESS_THAN_OR_EQUAL, 7)

    def test_data_validation_alert_style_values(self):
        """Test DataValidationAlertStyle enum values."""
        self.assertEqual(DataValidationAlertStyle.STOP, 0)
        self.assertEqual(DataValidationAlertStyle.WARNING, 1)
        self.assertEqual(DataValidationAlertStyle.INFORMATION, 2)

    def test_data_validation_ime_mode_values(self):
        """Test DataValidationImeMode enum values."""
        self.assertEqual(DataValidationImeMode.NO_CONTROL, 0)
        self.assertEqual(DataValidationImeMode.OFF, 1)
        self.assertEqual(DataValidationImeMode.ON, 2)
        self.assertEqual(DataValidationImeMode.DISABLED, 3)
        self.assertEqual(DataValidationImeMode.HIRAGANA, 4)
        self.assertEqual(DataValidationImeMode.FULL_KATAKANA, 5)
        self.assertEqual(DataValidationImeMode.HALF_KATAKANA, 6)
        self.assertEqual(DataValidationImeMode.FULL_ALPHA, 7)
        self.assertEqual(DataValidationImeMode.HALF_ALPHA, 8)
        self.assertEqual(DataValidationImeMode.FULL_HANGUL, 9)
        self.assertEqual(DataValidationImeMode.HALF_HANGUL, 10)


class TestDataValidation(unittest.TestCase):
    """Tests for DataValidation class."""

    def test_create_data_validation(self):
        """Test creating a DataValidation object."""
        dv = DataValidation("A1:A10")
        self.assertEqual(dv.sqref, "A1:A10")
        self.assertEqual(dv.type, DataValidationType.NONE)
        self.assertEqual(dv.operator, DataValidationOperator.BETWEEN)

    def test_data_validation_type_property(self):
        """Test setting validation type."""
        dv = DataValidation("A1:A10")
        dv.type = DataValidationType.WHOLE_NUMBER
        self.assertEqual(dv.type, DataValidationType.WHOLE_NUMBER)

        # Test setting with integer
        dv.type = 3
        self.assertEqual(dv.type, DataValidationType.LIST)

    def test_data_validation_operator_property(self):
        """Test setting validation operator."""
        dv = DataValidation("A1:A10")
        dv.operator = DataValidationOperator.GREATER_THAN
        self.assertEqual(dv.operator, DataValidationOperator.GREATER_THAN)

    def test_data_validation_formulas(self):
        """Test setting validation formulas."""
        dv = DataValidation("A1:A10")
        dv.formula1 = "1"
        dv.formula2 = "100"
        self.assertEqual(dv.formula1, "1")
        self.assertEqual(dv.formula2, "100")

    def test_data_validation_error_settings(self):
        """Test error message settings."""
        dv = DataValidation("A1:A10")
        dv.alert_style = DataValidationAlertStyle.WARNING
        dv.show_error_message = True
        dv.error_title = "Error"
        dv.error_message = "Please enter a valid value"

        self.assertEqual(dv.alert_style, DataValidationAlertStyle.WARNING)
        self.assertTrue(dv.show_error_message)
        self.assertEqual(dv.error_title, "Error")
        self.assertEqual(dv.error_message, "Please enter a valid value")

        # Test alias
        self.assertTrue(dv.show_error)
        self.assertEqual(dv.error, "Please enter a valid value")

    def test_data_validation_input_settings(self):
        """Test input message settings."""
        dv = DataValidation("A1:A10")
        dv.show_input_message = True
        dv.input_title = "Input"
        dv.input_message = "Enter a number"

        self.assertTrue(dv.show_input_message)
        self.assertEqual(dv.input_title, "Input")
        self.assertEqual(dv.input_message, "Enter a number")

        # Test aliases
        self.assertTrue(dv.show_input)
        self.assertEqual(dv.prompt_title, "Input")
        self.assertEqual(dv.prompt, "Enter a number")

    def test_data_validation_allow_blank(self):
        """Test allow blank setting."""
        dv = DataValidation("A1:A10")
        self.assertTrue(dv.allow_blank)  # Default is True

        dv.allow_blank = False
        self.assertFalse(dv.allow_blank)
        self.assertFalse(dv.ignore_blank)  # Alias

    def test_data_validation_show_dropdown(self):
        """Test show dropdown setting."""
        dv = DataValidation("A1:A10")
        self.assertTrue(dv.show_dropdown)  # Default is True

        dv.show_dropdown = False
        self.assertFalse(dv.show_dropdown)
        self.assertFalse(dv.in_cell_dropdown)  # Alias

    def test_data_validation_ime_mode(self):
        """Test IME mode setting."""
        dv = DataValidation("A1:A10")
        self.assertEqual(dv.ime_mode, DataValidationImeMode.NO_CONTROL)

        dv.ime_mode = DataValidationImeMode.HIRAGANA
        self.assertEqual(dv.ime_mode, DataValidationImeMode.HIRAGANA)

    def test_data_validation_add_method(self):
        """Test add() method for configuring validation."""
        dv = DataValidation("A1:A10")
        dv.add(DataValidationType.WHOLE_NUMBER,
               DataValidationAlertStyle.STOP,
               DataValidationOperator.BETWEEN,
               "1", "100")

        self.assertEqual(dv.type, DataValidationType.WHOLE_NUMBER)
        self.assertEqual(dv.alert_style, DataValidationAlertStyle.STOP)
        self.assertEqual(dv.operator, DataValidationOperator.BETWEEN)
        self.assertEqual(dv.formula1, "1")
        self.assertEqual(dv.formula2, "100")

    def test_data_validation_modify_method(self):
        """Test modify() method for changing validation."""
        dv = DataValidation("A1:A10")
        dv.type = DataValidationType.WHOLE_NUMBER
        dv.formula1 = "1"

        dv.modify(formula1="10", formula2="50")

        self.assertEqual(dv.type, DataValidationType.WHOLE_NUMBER)
        self.assertEqual(dv.formula1, "10")
        self.assertEqual(dv.formula2, "50")

    def test_data_validation_delete_method(self):
        """Test delete() method for clearing validation."""
        dv = DataValidation("A1:A10")
        dv.type = DataValidationType.WHOLE_NUMBER
        dv.formula1 = "1"
        dv.error_message = "Error"

        dv.delete()

        self.assertEqual(dv.type, DataValidationType.NONE)
        self.assertIsNone(dv.formula1)
        self.assertIsNone(dv.error_message)

    def test_data_validation_copy(self):
        """Test copying a DataValidation object."""
        dv = DataValidation("A1:A10")
        dv.type = DataValidationType.LIST
        dv.formula1 = '"Red,Green,Blue"'
        dv.error_message = "Select from list"

        dv_copy = dv.copy()

        self.assertEqual(dv_copy.sqref, "A1:A10")
        self.assertEqual(dv_copy.type, DataValidationType.LIST)
        self.assertEqual(dv_copy.formula1, '"Red,Green,Blue"')
        self.assertEqual(dv_copy.error_message, "Select from list")

        # Ensure it's a separate copy
        dv_copy.sqref = "B1:B10"
        self.assertEqual(dv.sqref, "A1:A10")

    def test_data_validation_string_truncation(self):
        """Test that strings are truncated to spec limits."""
        dv = DataValidation("A1:A10")

        # Error title max 32 chars
        dv.error_title = "A" * 50
        self.assertEqual(len(dv.error_title), 32)

        # Error message max 225 chars
        dv.error_message = "B" * 300
        self.assertEqual(len(dv.error_message), 225)

        # Input title max 32 chars
        dv.input_title = "C" * 50
        self.assertEqual(len(dv.input_title), 32)

        # Input message max 255 chars
        dv.input_message = "D" * 300
        self.assertEqual(len(dv.input_message), 255)


class TestDataValidationCollection(unittest.TestCase):
    """Tests for DataValidationCollection class."""

    def test_create_collection(self):
        """Test creating a DataValidationCollection."""
        collection = DataValidationCollection()
        self.assertEqual(collection.count, 0)
        self.assertEqual(len(collection), 0)

    def test_add_validation(self):
        """Test adding validations to collection."""
        collection = DataValidationCollection()

        dv = collection.add("A1:A10")
        self.assertEqual(collection.count, 1)
        self.assertEqual(dv.sqref, "A1:A10")

        dv2 = collection.add("B1:B10", DataValidationType.LIST, formula1='"Yes,No"')
        self.assertEqual(collection.count, 2)
        self.assertEqual(dv2.type, DataValidationType.LIST)

    def test_collection_iteration(self):
        """Test iterating over collection."""
        collection = DataValidationCollection()
        collection.add("A1:A10")
        collection.add("B1:B10")
        collection.add("C1:C10")

        refs = [dv.sqref for dv in collection]
        self.assertEqual(refs, ["A1:A10", "B1:B10", "C1:C10"])

    def test_collection_indexing(self):
        """Test indexing collection."""
        collection = DataValidationCollection()
        collection.add("A1:A10")
        collection.add("B1:B10")

        self.assertEqual(collection[0].sqref, "A1:A10")
        self.assertEqual(collection[1].sqref, "B1:B10")

    def test_remove_validation(self):
        """Test removing validation from collection."""
        collection = DataValidationCollection()
        dv1 = collection.add("A1:A10")
        dv2 = collection.add("B1:B10")

        result = collection.remove(dv1)
        self.assertTrue(result)
        self.assertEqual(collection.count, 1)
        self.assertEqual(collection[0].sqref, "B1:B10")

    def test_remove_at(self):
        """Test removing validation at index."""
        collection = DataValidationCollection()
        collection.add("A1:A10")
        collection.add("B1:B10")
        collection.add("C1:C10")

        collection.remove_at(1)
        self.assertEqual(collection.count, 2)
        self.assertEqual(collection[0].sqref, "A1:A10")
        self.assertEqual(collection[1].sqref, "C1:C10")

    def test_clear_collection(self):
        """Test clearing collection."""
        collection = DataValidationCollection()
        collection.add("A1:A10")
        collection.add("B1:B10")

        collection.clear()
        self.assertEqual(collection.count, 0)


class TestDataValidationWorksheetIntegration(unittest.TestCase):
    """Tests for data validation integration with Worksheet."""

    def test_worksheet_has_data_validations(self):
        """Test that worksheet has data_validations property."""
        wb = Workbook()
        ws = wb.worksheets[0]

        self.assertIsNotNone(ws.data_validations)
        self.assertIsInstance(ws.data_validations, DataValidationCollection)
        self.assertEqual(ws.data_validations.count, 0)

    def test_add_validation_to_worksheet(self):
        """Test adding validation to worksheet."""
        wb = Workbook()
        ws = wb.worksheets[0]

        dv = ws.data_validations.add("A1:A10")
        dv.type = DataValidationType.WHOLE_NUMBER
        dv.operator = DataValidationOperator.BETWEEN
        dv.formula1 = "1"
        dv.formula2 = "100"

        self.assertEqual(ws.data_validations.count, 1)
        self.assertEqual(ws.data_validations[0].type, DataValidationType.WHOLE_NUMBER)


class TestDataValidationRoundtrip(unittest.TestCase):
    """Tests for saving and loading data validations."""

    def setUp(self):
        """Set up test fixtures."""
        # Use outputfiles directory for persistent test files
        self.output_dir = os.path.join(os.path.dirname(__file__), "outputfiles")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def test_whole_number_validation_roundtrip(self):
        """Test whole number validation save and load."""
        # Create workbook with validation
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers and sample data
        ws.cells['A1'].value = "Data Validation: Whole Number (1-100)"
        ws.cells['A2'].value = "Range: A3:A10"
        ws.cells['A3'].value = "Try entering: 50 (valid)"
        ws.cells['A4'].value = "Try entering: 150 (invalid)"
        ws.cells['A5'].value = "Try entering: 0 (invalid)"
        ws.cells['A6'].value = "Try entering: -5 (invalid)"
        ws.cells['A7'].value = "Try entering: 100 (valid)"

        dv = ws.data_validations.add("A3:A10")
        dv.type = DataValidationType.WHOLE_NUMBER
        dv.operator = DataValidationOperator.BETWEEN
        dv.formula1 = "1"
        dv.formula2 = "100"
        dv.allow_blank = True
        dv.show_error_message = True
        dv.error_title = "Invalid"
        dv.error_message = "Enter 1-100"
        dv.show_input_message = True
        dv.input_title = "Input"
        dv.input_message = "Enter a number"

        # Save
        file_path = os.path.join(self.output_dir, "test_whole_number.xlsx")
        wb.save(file_path)

        # Load and verify
        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        self.assertEqual(ws2.data_validations.count, 1)
        dv2 = ws2.data_validations[0]

        self.assertEqual(dv2.sqref, "A3:A10")
        self.assertEqual(dv2.type, DataValidationType.WHOLE_NUMBER)
        self.assertEqual(dv2.operator, DataValidationOperator.BETWEEN)
        self.assertEqual(dv2.formula1, "1")
        self.assertEqual(dv2.formula2, "100")
        self.assertTrue(dv2.allow_blank)
        self.assertTrue(dv2.show_error_message)
        self.assertEqual(dv2.error_title, "Invalid")
        self.assertEqual(dv2.error_message, "Enter 1-100")

    def test_list_validation_roundtrip(self):
        """Test list validation save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "Data Validation: List (Dropdown)"
        ws.cells['A2'].value = "Range: B3:B10"
        ws.cells['A3'].value = "Click cell B3 to see dropdown options:"
        ws.cells['B3'].value = "Red"
        ws.cells['A4'].value = "Options: Red, Green, Blue"

        dv = ws.data_validations.add("B3:B10")
        dv.type = DataValidationType.LIST
        dv.formula1 = '"Red,Green,Blue"'
        dv.show_dropdown = True
        dv.show_error_message = True
        dv.error_message = "Select from list"

        file_path = os.path.join(self.output_dir, "test_list.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        self.assertEqual(ws2.data_validations.count, 1)
        dv2 = ws2.data_validations[0]

        self.assertEqual(dv2.type, DataValidationType.LIST)
        self.assertEqual(dv2.sqref, "B3:B10")
        self.assertEqual(dv2.formula1, '"Red,Green,Blue"')
        self.assertTrue(dv2.show_dropdown)

    def test_date_validation_roundtrip(self):
        """Test date validation save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "Data Validation: Date (>= 2023-01-01)"
        ws.cells['A2'].value = "Range: C3:C10"
        ws.cells['A3'].value = "Try entering: 2023-06-15 (valid)"
        ws.cells['A4'].value = "Try entering: 2022-12-31 (invalid - warning)"
        ws.cells['A5'].value = "Alert Style: WARNING"

        dv = ws.data_validations.add("C3:C10")
        dv.type = DataValidationType.DATE
        dv.operator = DataValidationOperator.GREATER_THAN_OR_EQUAL
        dv.formula1 = "44927"  # 2023-01-01 in Excel serial format
        dv.alert_style = DataValidationAlertStyle.WARNING

        file_path = os.path.join(self.output_dir, "test_date.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        dv2 = ws2.data_validations[0]
        self.assertEqual(dv2.type, DataValidationType.DATE)
        self.assertEqual(dv2.sqref, "C3:C10")
        self.assertEqual(dv2.operator, DataValidationOperator.GREATER_THAN_OR_EQUAL)
        self.assertEqual(dv2.alert_style, DataValidationAlertStyle.WARNING)

    def test_custom_validation_roundtrip(self):
        """Test custom formula validation save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "Data Validation: Custom Formula"
        ws.cells['A2'].value = "Range: D3:D10"
        ws.cells['A3'].value = "Formula: =AND(D3>0,D3<1000)"
        ws.cells['A4'].value = "Valid: 1-999"
        ws.cells['A5'].value = "Invalid: 0, 1000, negative numbers"

        dv = ws.data_validations.add("D3:D10")
        dv.type = DataValidationType.CUSTOM
        dv.formula1 = "=AND(D3>0,D3<1000)"

        file_path = os.path.join(self.output_dir, "test_custom.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        dv2 = ws2.data_validations[0]
        self.assertEqual(dv2.type, DataValidationType.CUSTOM)
        self.assertEqual(dv2.sqref, "D3:D10")
        self.assertEqual(dv2.formula1, "=AND(D3>0,D3<1000)")

    def test_multiple_validations_roundtrip(self):
        """Test multiple validations save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "Multiple Data Validations Demo"
        ws.cells['A2'].value = "Column A: Whole Number (> 0)"
        ws.cells['A3'].value = "Try: 10 (valid), -5 (invalid)"
        ws.cells['B2'].value = "Column B: List (A,B,C)"
        ws.cells['B3'].value = "Click for dropdown"
        ws.cells['C2'].value = "Column C: Text Length (<= 50)"
        ws.cells['C3'].value = "Try: short (valid), very long text (invalid)"

        # Add multiple validations
        dv1 = ws.data_validations.add("A3:A10")
        dv1.type = DataValidationType.WHOLE_NUMBER
        dv1.operator = DataValidationOperator.GREATER_THAN
        dv1.formula1 = "0"

        dv2 = ws.data_validations.add("B3:B10")
        dv2.type = DataValidationType.LIST
        dv2.formula1 = '"A,B,C"'

        dv3 = ws.data_validations.add("C3:C10")
        dv3.type = DataValidationType.TEXT_LENGTH
        dv3.operator = DataValidationOperator.LESS_THAN_OR_EQUAL
        dv3.formula1 = "50"

        file_path = os.path.join(self.output_dir, "test_multiple.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        self.assertEqual(ws2.data_validations.count, 3)

        # Verify each validation
        refs = [dv.sqref for dv in ws2.data_validations]
        self.assertIn("A3:A10", refs)
        self.assertIn("B3:B10", refs)
        self.assertIn("C3:C10", refs)

    def test_text_length_validation_roundtrip(self):
        """Test text length validation save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "Data Validation: Text Length (5-50 chars)"
        ws.cells['A2'].value = "Range: E3:E10"
        ws.cells['A3'].value = "Valid: 'Hello' (5 chars)"
        ws.cells['A4'].value = "Valid: 'This is a medium length text' (30 chars)"
        ws.cells['A5'].value = "Invalid: 'Hi' (2 chars)"
        ws.cells['A6'].value = "Invalid: 'This text is way too long and exceeds fifty characters limit' (60 chars)"

        dv = ws.data_validations.add("E3:E10")
        dv.type = DataValidationType.TEXT_LENGTH
        dv.operator = DataValidationOperator.BETWEEN
        dv.formula1 = "5"
        dv.formula2 = "50"

        file_path = os.path.join(self.output_dir, "test_textlength.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        dv2 = ws2.data_validations[0]
        self.assertEqual(dv2.type, DataValidationType.TEXT_LENGTH)
        self.assertEqual(dv2.sqref, "E3:E10")
        self.assertEqual(dv2.formula1, "5")
        self.assertEqual(dv2.formula2, "50")

    def test_decimal_validation_roundtrip(self):
        """Test decimal validation save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "Data Validation: Decimal (NOT BETWEEN 0 and 1)"
        ws.cells['A2'].value = "Range: F3:F10"
        ws.cells['A3'].value = "Valid: -5.5, 1.5, 100.25"
        ws.cells['A4'].value = "Invalid: 0, 0.5, 0.99, 1.0"
        ws.cells['A5'].value = "Operator: NOT_BETWEEN"

        dv = ws.data_validations.add("F3:F10")
        dv.type = DataValidationType.DECIMAL
        dv.operator = DataValidationOperator.NOT_BETWEEN
        dv.formula1 = "0"
        dv.formula2 = "1"

        file_path = os.path.join(self.output_dir, "test_decimal.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        dv2 = ws2.data_validations[0]
        self.assertEqual(dv2.type, DataValidationType.DECIMAL)
        self.assertEqual(dv2.sqref, "F3:F10")
        self.assertEqual(dv2.operator, DataValidationOperator.NOT_BETWEEN)

    def test_time_validation_roundtrip(self):
        """Test time validation save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "Data Validation: Time (9:00 AM - 5:00 PM)"
        ws.cells['A2'].value = "Range: G3:G10"
        ws.cells['A3'].value = "Valid: 9:00 AM, 12:00 PM, 5:00 PM"
        ws.cells['A4'].value = "Invalid: 8:59 AM, 5:01 PM, 12:00 AM"
        ws.cells['A5'].value = "Format: Excel time (0.375 = 9:00 AM, 0.708 = 5:00 PM)"

        dv = ws.data_validations.add("G3:G10")
        dv.type = DataValidationType.TIME
        dv.operator = DataValidationOperator.BETWEEN
        dv.formula1 = "0.375"  # 9:00 AM
        dv.formula2 = "0.708"  # 5:00 PM

        file_path = os.path.join(self.output_dir, "test_time.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        dv2 = ws2.data_validations[0]
        self.assertEqual(dv2.type, DataValidationType.TIME)
        self.assertEqual(dv2.sqref, "G3:G10")

    def test_all_operators_roundtrip(self):
        """Test all comparison operators save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "All 8 Data Validation Operators"
        ws.cells['A2'].value = "Column A: BETWEEN (10-100)"
        ws.cells['A3'].value = "Try: 50 (valid), 5 (invalid)"
        ws.cells['B2'].value = "Column B: NOT_BETWEEN (10-100)"
        ws.cells['B3'].value = "Try: 5 (valid), 50 (invalid)"
        ws.cells['C2'].value = "Column C: EQUAL (=10)"
        ws.cells['C3'].value = "Try: 10 (valid), 11 (invalid)"
        ws.cells['D2'].value = "Column D: NOT_EQUAL (<>10)"
        ws.cells['D3'].value = "Try: 11 (valid), 10 (invalid)"
        ws.cells['E2'].value = "Column E: GREATER_THAN (>10)"
        ws.cells['E3'].value = "Try: 11 (valid), 10 (invalid)"
        ws.cells['F2'].value = "Column F: LESS_THAN (<10)"
        ws.cells['F3'].value = "Try: 9 (valid), 10 (invalid)"
        ws.cells['G2'].value = "Column G: GREATER_THAN_OR_EQUAL (>=10)"
        ws.cells['G3'].value = "Try: 10 (valid), 9 (invalid)"
        ws.cells['H2'].value = "Column H: LESS_THAN_OR_EQUAL (<=10)"
        ws.cells['H3'].value = "Try: 10 (valid), 11 (invalid)"

        operators = [
            (DataValidationOperator.BETWEEN, "A4:A10"),
            (DataValidationOperator.NOT_BETWEEN, "B4:B10"),
            (DataValidationOperator.EQUAL, "C4:C10"),
            (DataValidationOperator.NOT_EQUAL, "D4:D10"),
            (DataValidationOperator.GREATER_THAN, "E4:E10"),
            (DataValidationOperator.LESS_THAN, "F4:F10"),
            (DataValidationOperator.GREATER_THAN_OR_EQUAL, "G4:G10"),
            (DataValidationOperator.LESS_THAN_OR_EQUAL, "H4:H10"),
        ]

        for op, sqref in operators:
            dv = ws.data_validations.add(sqref)
            dv.type = DataValidationType.WHOLE_NUMBER
            dv.operator = op
            dv.formula1 = "10"
            if op in (DataValidationOperator.BETWEEN, DataValidationOperator.NOT_BETWEEN):
                dv.formula2 = "100"

        file_path = os.path.join(self.output_dir, "test_all_operators.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        self.assertEqual(ws2.data_validations.count, 8)

        # Create a map of sqref to validation
        dv_map = {dv.sqref: dv for dv in ws2.data_validations}

        # Updated sqrefs to match new ranges
        updated_sqrefs = [
            ("A4:A10", DataValidationOperator.BETWEEN),
            ("B4:B10", DataValidationOperator.NOT_BETWEEN),
            ("C4:C10", DataValidationOperator.EQUAL),
            ("D4:D10", DataValidationOperator.NOT_EQUAL),
            ("E4:E10", DataValidationOperator.GREATER_THAN),
            ("F4:F10", DataValidationOperator.LESS_THAN),
            ("G4:G10", DataValidationOperator.GREATER_THAN_OR_EQUAL),
            ("H4:H10", DataValidationOperator.LESS_THAN_OR_EQUAL),
        ]

        for sqref, op in updated_sqrefs:
            self.assertEqual(dv_map[sqref].operator, op)

    def test_alert_styles_roundtrip(self):
        """Test all alert styles save and load."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add descriptive headers
        ws.cells['A1'].value = "All 3 Data Validation Alert Styles"
        ws.cells['A2'].value = "Column A: STOP (blocks invalid entry)"
        ws.cells['A3'].value = "Try: 0 (blocked), must retry"
        ws.cells['B2'].value = "Column B: WARNING (allows with warning)"
        ws.cells['B3'].value = "Try: 0 (allowed with warning)"
        ws.cells['C2'].value = "Column C: INFORMATION (allows with info)"
        ws.cells['C3'].value = "Try: 0 (allowed with info message)"

        styles = [
            (DataValidationAlertStyle.STOP, "A4:A10"),
            (DataValidationAlertStyle.WARNING, "B4:B10"),
            (DataValidationAlertStyle.INFORMATION, "C4:C10"),
        ]

        for style, sqref in styles:
            dv = ws.data_validations.add(sqref)
            dv.type = DataValidationType.WHOLE_NUMBER
            dv.alert_style = style
            dv.formula1 = "1"
            dv.formula2 = "10"

        file_path = os.path.join(self.output_dir, "test_alert_styles.xlsx")
        wb.save(file_path)

        wb2 = Workbook(file_path)
        ws2 = wb2.worksheets[0]

        dv_map = {dv.sqref: dv for dv in ws2.data_validations}

        # Updated sqrefs to match new ranges
        updated_styles = [
            ("A4:A10", DataValidationAlertStyle.STOP),
            ("B4:B10", DataValidationAlertStyle.WARNING),
            ("C4:C10", DataValidationAlertStyle.INFORMATION),
        ]

        for sqref, style in updated_styles:
            self.assertEqual(dv_map[sqref].alert_style, style)


if __name__ == '__main__':
    unittest.main(verbosity=2)
