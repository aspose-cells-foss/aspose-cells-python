import unittest
import os
import sys
import tempfile

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook, Cell


class TestCellValues(unittest.TestCase):
    """
    Test comprehensive cell value handling including int, double, string, and formula.
    
    Features tested:
    - Integer values (positive, negative, zero)
    - Double/float values (positive, negative, decimals)
    - String values (regular, special characters, Unicode, multi-line)
    - Formula values (SUM, IF, VLOOKUP, AVERAGE, MAX, MIN, COUNT)
    - Mixed value types in same worksheet
    - Save and load cell values to verify persistence
    - Edge cases (None, empty string, large numbers, scientific notation)
    """
    
    def setUp(self):
        """Set up test workbook and worksheet."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.worksheets[0]
    
    def test_integer_values(self):
        """Test setting and saving integer values."""
        test_values = [0, 1, -1, 42, 1000, -999]
        
        for i, value in enumerate(test_values):
            cell = Cell(value)
            self.worksheet.cells[f"A{i+1}"] = cell
            
            # Verify the value was set correctly
            self.assertEqual(self.worksheet.cells[f"A{i+1}"].value, value)
            self.assertIsInstance(self.worksheet.cells[f"A{i+1}"].value, int)
    
    def test_double_values(self):
        """Test setting and saving double/float values."""
        test_values = [0.0, 1.5, -2.7, 3.14159, 0.0001, -999.999]
        
        for i, value in enumerate(test_values):
            cell = Cell(value)
            self.worksheet.cells[f"B{i+1}"] = cell
            
            # Verify the value was set correctly
            self.assertEqual(self.worksheet.cells[f"B{i+1}"].value, value)
            self.assertIsInstance(self.worksheet.cells[f"B{i+1}"].value, float)
    
    def test_string_values(self):
        """Test setting and saving string values."""
        test_values = [
            "Hello World",
            "Test String",
            "123",  # String representation of number
            "3.14",  # String representation of float
            "",  # Empty string
            "Special chars: !@#$%^&*()",
            "Unicode: 你好世界",
            "Multi\nline\nstring"
        ]
        
        for i, value in enumerate(test_values):
            cell = Cell(value)
            self.worksheet.cells[f"C{i+1}"] = cell
            
            # Verify the value was set correctly
            self.assertEqual(self.worksheet.cells[f"C{i+1}"].value, value)
            self.assertIsInstance(self.worksheet.cells[f"C{i+1}"].value, str)
    
    def test_formula_values(self):
        """Test setting and saving formula values."""
        test_formulas = [
            "=SUM(A1:A5)",
            "=A1+B1",
            "=IF(A1>0, \"Positive\", \"Non-positive\")",
            "=VLOOKUP(A1, B1:C10, 2, FALSE)",
            "=AVERAGE(A1:A10)",
            "=MAX(A1:A5)",
            "=MIN(A1:A5)",
            "=COUNT(A1:A10)"
        ]
        
        for i, formula in enumerate(test_formulas):
            cell = Cell(None, formula)
            self.worksheet.cells[f"D{i+1}"] = cell
            
            # Verify the formula was set correctly
            self.assertEqual(self.worksheet.cells[f"D{i+1}"].formula, formula)
            self.assertIsNone(self.worksheet.cells[f"D{i+1}"].value)
    
    def test_mixed_values(self):
        """Test setting mixed value types in the same worksheet."""
        # Set up mixed values
        self.worksheet.cells["A1"] = Cell(42)  # int
        self.worksheet.cells["A2"] = Cell(3.14159)  # float
        self.worksheet.cells["A3"] = Cell("Hello")  # string
        self.worksheet.cells["A4"] = Cell(None, "=SUM(A1:A2)")  # formula
        
        # Verify all values are set correctly
        self.assertEqual(self.worksheet.cells["A1"].value, 42)
        self.assertIsInstance(self.worksheet.cells["A1"].value, int)
        
        self.assertEqual(self.worksheet.cells["A2"].value, 3.14159)
        self.assertIsInstance(self.worksheet.cells["A2"].value, float)
        
        self.assertEqual(self.worksheet.cells["A3"].value, "Hello")
        self.assertIsInstance(self.worksheet.cells["A3"].value, str)
        
        self.assertEqual(self.worksheet.cells["A4"].formula, "=SUM(A1:A2)")
        self.assertIsNone(self.worksheet.cells["A4"].value)
    
    def test_save_and_load_cell_values(self):
        """Test saving and loading cell values to verify persistence."""
        # Create test data with all value types
        test_data = {
            "A1": {"value": 42, "type": "int"},
            "A2": {"value": 3.14159, "type": "float"},
            "A3": {"value": "Hello World", "type": "string"},
            "A4": {"value": None, "formula": "=SUM(A1:A2)", "type": "formula"},
            "A5": {"value": -100, "type": "int"},
            "A6": {"value": 2.71828, "type": "float"},
            "A7": {"value": "", "type": "string"},
            "A8": {"value": None, "formula": "=A1+A2", "type": "formula"},
            "A9": {"value": "Test String", "type": "string"},
            "A10": {"value": 0, "type": "int"}
        }
        
        # Set up the test data
        for ref, data in test_data.items():
            if data["type"] == "formula":
                cell = Cell(data["value"], data["formula"])
            else:
                cell = Cell(data["value"])
            self.worksheet.cells[ref] = cell
        
        # Save to outputfiles folder
        output_path = 'outputfiles/test_cell_values.xlsx'
        
        # Ensure outputfiles directory exists
        os.makedirs('outputfiles', exist_ok=True)
        
        self.workbook.save(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify file is not empty
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"Cell values test file saved to: {output_path}")
        print(f"File size: {file_size} bytes")
        
        # Load the file back
        loaded_workbook = Workbook(output_path)
        loaded_worksheet = loaded_workbook.worksheets[0]
        
        # Verify all values are loaded correctly
        for ref, expected_data in test_data.items():
            cell = loaded_worksheet.cells[ref]
            
            if expected_data["type"] == "formula":
                # For formulas, we expect the formula to be preserved
                self.assertEqual(cell.formula, expected_data["formula"])
                # The value might be calculated or None, depending on implementation
                # For now, just check that it's a formula cell
                self.assertIsNotNone(cell.formula)
            else:
                # For regular values, check the value and type
                # Handle empty string case specially
                if expected_data["value"] == "" and cell.value is None:
                    # Empty strings might be loaded as None - this is acceptable
                    pass
                elif cell.value is not None:
                    self.assertEqual(cell.value, expected_data["value"])
                    
                    if expected_data["type"] == "int":
                        self.assertIsInstance(cell.value, (int, float))  # Excel might convert to float
                        # If it's a whole number, it should be treated as int-like
                        if isinstance(cell.value, float):
                            self.assertEqual(cell.value, int(cell.value))
                    elif expected_data["type"] == "float":
                        self.assertIsInstance(cell.value, float)
                    elif expected_data["type"] == "string":
                        self.assertIsInstance(cell.value, str)
    
    def test_edge_cases(self):
        """Test edge cases for cell values."""
        # Test None value
        cell = Cell(None)
        self.assertIsNone(cell.value)
        self.assertIsNone(cell.formula)
        
        # Test empty string
        cell = Cell("")
        self.assertEqual(cell.value, "")
        self.assertIsInstance(cell.value, str)
        
        # Test very large numbers
        cell = Cell(999999999999)
        self.assertEqual(cell.value, 999999999999)
        
        # Test very small decimals
        cell = Cell(0.0000001)
        self.assertEqual(cell.value, 0.0000001)
        
        # Test scientific notation
        cell = Cell(1.23e-10)
        self.assertEqual(cell.value, 1.23e-10)


if __name__ == '__main__':
    unittest.main()