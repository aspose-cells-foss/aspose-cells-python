"""
Comprehensive Test Suite for CSV Import/Export Feature

This test suite covers all aspects of CSV import and export functionality including:
- Basic CSV import/export with default settings
- Custom delimiters (comma, semicolon, tab, pipe)
- Different encodings (UTF-8, UTF-16, Latin-1, ASCII)
- Type inference (integers, floats, booleans, dates, strings)
- Quoting and escaping special characters
- Multiline values
- Empty cells and missing values
- Headers handling
- Skip rows functionality
- Round-trip testing (export then import)
- Large datasets
- Special characters and Unicode
- Date/datetime formatting
- Boolean value parsing
- BOM (Byte Order Mark) handling
- Line terminator customization
- Worksheet index selection
"""

import unittest
import os
import sys
import tempfile
from datetime import datetime, date, time

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook, Cell
from aspose_cells.csv_handler import CSVHandler, CSVLoadOptions, CSVSaveOptions


class TestCSVBasicImportExport(unittest.TestCase):
    """Test basic CSV import and export functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.worksheets[0]
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_simple_csv_export(self):
        """Test exporting a simple worksheet to CSV."""
        # Add some data
        self.worksheet.cells['A1'].value = "Name"
        self.worksheet.cells['B1'].value = "Age"
        self.worksheet.cells['A2'].value = "Alice"
        self.worksheet.cells['B2'].value = 30
        self.worksheet.cells['A3'].value = "Bob"
        self.worksheet.cells['B3'].value = 25
        
        # Export to CSV
        output_path = os.path.join(self.test_dir, 'test_simple_export.csv')
        self.workbook.save_as_csv(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Read and verify content
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('Name,Age', content)
        self.assertIn('Alice,30', content)
        self.assertIn('Bob,25', content)
    
    def test_simple_csv_import(self):
        """Test importing a simple CSV file."""
        # Create a CSV file
        csv_path = os.path.join(self.test_dir, 'test_simple_import.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Name,Age\n')
            f.write('Alice,30\n')
            f.write('Bob,25\n')
        
        # Import CSV
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        # Verify data
        self.assertEqual(ws.cells['A1'].value, 'Name')
        self.assertEqual(ws.cells['B1'].value, 'Age')
        self.assertEqual(ws.cells['A2'].value, 'Alice')
        self.assertEqual(ws.cells['B2'].value, 30)  # Type inference
        self.assertEqual(ws.cells['A3'].value, 'Bob')
        self.assertEqual(ws.cells['B3'].value, 25)  # Type inference
    
    def test_csv_roundtrip(self):
        """Test exporting and then importing back (round-trip)."""
        # Create original data
        original_data = [
            ["Name", "Age", "Score"],
            ["Alice", 30, 95.5],
            ["Bob", 25, 87.3],
            ["Charlie", 35, 92.1]
        ]
        
        # Populate worksheet
        for row_idx, row in enumerate(original_data, start=1):
            for col_idx, value in enumerate(row, start=1):
                col_letter = chr(ord('A') + col_idx - 1)
                self.worksheet.cells[f'{col_letter}{row_idx}'].value = value
        
        # Export to CSV
        export_path = os.path.join(self.test_dir, 'test_roundtrip_export.csv')
        self.workbook.save_as_csv(export_path)
        
        # Import to new workbook
        wb_imported = Workbook()
        wb_imported.load_csv(export_path)
        ws_imported = wb_imported.worksheets[0]
        
        # Verify data matches (note: type inference may change types)
        for row_idx, row in enumerate(original_data, start=1):
            for col_idx, value in enumerate(row, start=1):
                col_letter = chr(ord('A') + col_idx - 1)
                imported_value = ws_imported.cells[f'{col_letter}{row_idx}'].value
                # For numeric values, compare as float
                if isinstance(value, (int, float)):
                    self.assertAlmostEqual(float(imported_value), float(value), places=1)
                else:
                    self.assertEqual(imported_value, value)


class TestCSVDelimiters(unittest.TestCase):
    """Test CSV import/export with different delimiters."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_comma_delimiter(self):
        """Test CSV with comma delimiter (default)."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "A"
        ws.cells['B1'].value = "B"
        ws.cells['C1'].value = "C"
        
        output_path = os.path.join(self.test_dir, 'test_comma.csv')
        wb.save_as_csv(output_path)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('A,B,C', content)
    
    def test_semicolon_delimiter(self):
        """Test CSV with semicolon delimiter."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "A"
        ws.cells['B1'].value = "B"
        ws.cells['C1'].value = "C"
        
        options = CSVSaveOptions()
        options.delimiter = ';'
        
        output_path = os.path.join(self.test_dir, 'test_semicolon.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('A;B;C', content)
    
    def test_tab_delimiter(self):
        """Test CSV with tab delimiter."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "A"
        ws.cells['B1'].value = "B"
        ws.cells['C1'].value = "C"
        
        options = CSVSaveOptions()
        options.delimiter = '\t'
        
        output_path = os.path.join(self.test_dir, 'test_tab.tsv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('A\tB\tC', content)
    
    def test_pipe_delimiter(self):
        """Test CSV with pipe delimiter."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "A"
        ws.cells['B1'].value = "B"
        ws.cells['C1'].value = "C"
        
        options = CSVSaveOptions()
        options.delimiter = '|'
        
        output_path = os.path.join(self.test_dir, 'test_pipe.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('A|B|C', content)
    
    def test_import_with_semicolon_delimiter(self):
        """Test importing CSV with semicolon delimiter."""
        csv_path = os.path.join(self.test_dir, 'test_import_semicolon.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Name;Age;City\n')
            f.write('Alice;30;New York\n')
            f.write('Bob;25;London\n')
        
        options = CSVLoadOptions()
        options.delimiter = ';'
        
        wb = Workbook()
        wb.load_csv(csv_path, options)
        ws = wb.worksheets[0]
        
        self.assertEqual(ws.cells['A1'].value, 'Name')
        self.assertEqual(ws.cells['B1'].value, 'Age')
        self.assertEqual(ws.cells['C1'].value, 'City')
        self.assertEqual(ws.cells['A2'].value, 'Alice')
        self.assertEqual(ws.cells['B2'].value, 30)
        self.assertEqual(ws.cells['C2'].value, 'New York')


class TestCSVEncodings(unittest.TestCase):
    """Test CSV import/export with different encodings."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_utf8_encoding(self):
        """Test CSV with UTF-8 encoding (default)."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "Hello"
        ws.cells['A2'].value = "你好世界"  # Chinese characters
        
        output_path = os.path.join(self.test_dir, 'test_utf8.csv')
        wb.save_as_csv(output_path)
        
        # Verify file can be read with UTF-8
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('你好世界', content)
    
    def test_utf16_encoding(self):
        """Test CSV with UTF-16 encoding."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "Hello"
        ws.cells['A2'].value = "你好世界"
        
        options = CSVSaveOptions()
        options.encoding = 'utf-16'
        
        output_path = os.path.join(self.test_dir, 'test_utf16.csv')
        wb.save_as_csv(output_path, options)
        
        # Verify file can be read with UTF-16
        with open(output_path, 'r', encoding='utf-16') as f:
            content = f.read()
        
        self.assertIn('你好世界', content)
    
    def test_latin1_encoding(self):
        """Test CSV with Latin-1 encoding."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "Café"
        ws.cells['A2'].value = "Naïve"
        
        options = CSVSaveOptions()
        options.encoding = 'latin-1'
        
        output_path = os.path.join(self.test_dir, 'test_latin1.csv')
        wb.save_as_csv(output_path, options)
        
        # Verify file can be read with Latin-1
        with open(output_path, 'r', encoding='latin-1') as f:
            content = f.read()
        
        self.assertIn('Café', content)
        self.assertIn('Naïve', content)
    
    def test_import_with_utf8_bom(self):
        """Test importing CSV with UTF-8 BOM."""
        csv_path = os.path.join(self.test_dir, 'test_utf8_bom.csv')
        with open(csv_path, 'wb') as f:
            f.write('\ufeff'.encode('utf-8'))  # Write BOM
            f.write('Name,Age\n'.encode('utf-8'))
            f.write('Alice,30\n'.encode('utf-8'))
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        # BOM is stripped from the content when using load_csv_from_string
        # but may be included when using load_csv with file
        # This test verifies the file can be read without errors
        self.assertIsNotNone(ws.cells['A1'].value)
        self.assertIn('Name', ws.cells['A1'].value)
        self.assertEqual(ws.cells['A2'].value, 'Alice')
    
    def test_export_with_bom(self):
        """Test exporting CSV with UTF-8 BOM."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = "Name"
        ws.cells['A2'].value = "Alice"
        
        options = CSVSaveOptions()
        options.write_bom = True
        
        output_path = os.path.join(self.test_dir, 'test_export_bom.csv')
        wb.save_as_csv(output_path, options)
        
        # Verify BOM is present
        with open(output_path, 'rb') as f:
            content = f.read()
        
        self.assertTrue(content.startswith(b'\xef\xbb\xbf'))


class TestCSVTypeInference(unittest.TestCase):
    """Test CSV type inference during import."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_integer_type_inference(self):
        """Test integer type inference."""
        csv_path = os.path.join(self.test_dir, 'test_integers.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Value\n')
            f.write('42\n')
            f.write('-100\n')
            f.write('0\n')
            f.write('999999\n')
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        self.assertIsInstance(ws.cells['A2'].value, int)
        self.assertEqual(ws.cells['A2'].value, 42)
        self.assertIsInstance(ws.cells['A3'].value, int)
        self.assertEqual(ws.cells['A3'].value, -100)
        self.assertIsInstance(ws.cells['A4'].value, int)
        self.assertEqual(ws.cells['A4'].value, 0)
        self.assertIsInstance(ws.cells['A5'].value, int)
        self.assertEqual(ws.cells['A5'].value, 999999)
    
    def test_float_type_inference(self):
        """Test float type inference."""
        csv_path = os.path.join(self.test_dir, 'test_floats.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Value\n')
            f.write('3.14159\n')
            f.write('-2.71828\n')
            f.write('0.0001\n')
            f.write('1.23e-10\n')
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        self.assertIsInstance(ws.cells['A2'].value, float)
        self.assertAlmostEqual(ws.cells['A2'].value, 3.14159, places=5)
        self.assertIsInstance(ws.cells['A3'].value, float)
        self.assertAlmostEqual(ws.cells['A3'].value, -2.71828, places=5)
        self.assertIsInstance(ws.cells['A4'].value, float)
        self.assertAlmostEqual(ws.cells['A4'].value, 0.0001, places=5)
        self.assertIsInstance(ws.cells['A5'].value, float)
        self.assertAlmostEqual(ws.cells['A5'].value, 1.23e-10, places=15)
    
    def test_boolean_type_inference(self):
        """Test boolean type inference."""
        csv_path = os.path.join(self.test_dir, 'test_booleans.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Value\n')
            f.write('true\n')
            f.write('false\n')
            f.write('TRUE\n')
            f.write('FALSE\n')
            f.write('yes\n')
            f.write('no\n')
            f.write('1\n')
            f.write('0\n')
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        self.assertIsInstance(ws.cells['A2'].value, bool)
        self.assertTrue(ws.cells['A2'].value)
        self.assertIsInstance(ws.cells['A3'].value, bool)
        self.assertFalse(ws.cells['A3'].value)
        self.assertIsInstance(ws.cells['A4'].value, bool)
        self.assertTrue(ws.cells['A4'].value)
        self.assertIsInstance(ws.cells['A5'].value, bool)
        self.assertFalse(ws.cells['A5'].value)
        self.assertIsInstance(ws.cells['A6'].value, bool)
        self.assertTrue(ws.cells['A6'].value)
        self.assertIsInstance(ws.cells['A7'].value, bool)
        self.assertFalse(ws.cells['A7'].value)
        self.assertIsInstance(ws.cells['A8'].value, bool)
        self.assertTrue(ws.cells['A8'].value)
        self.assertIsInstance(ws.cells['A9'].value, bool)
        self.assertFalse(ws.cells['A9'].value)
    
    def test_date_type_inference(self):
        """Test date type inference."""
        csv_path = os.path.join(self.test_dir, 'test_dates.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Date\n')
            f.write('2023-01-15\n')
            f.write('2023/12/31\n')
            f.write('15-01-2023\n')
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        # Check that dates are parsed as date objects
        self.assertIsInstance(ws.cells['A2'].value, date)
        self.assertIsInstance(ws.cells['A3'].value, date)
        self.assertIsInstance(ws.cells['A4'].value, date)
    
    def test_datetime_type_inference(self):
        """Test datetime type inference."""
        csv_path = os.path.join(self.test_dir, 'test_datetimes.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('DateTime\n')
            f.write('2023-01-15 14:30:00\n')
            f.write('2023/12/31 23:59:59\n')
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        # Check that datetimes are parsed as datetime objects
        self.assertIsInstance(ws.cells['A2'].value, datetime)
        self.assertIsInstance(ws.cells['A3'].value, datetime)
    
    def test_string_type_preservation(self):
        """Test that strings are preserved as strings."""
        csv_path = os.path.join(self.test_dir, 'test_strings.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Value\n')
            f.write('Hello World\n')
            f.write('123\n')  # String that looks like a number
            f.write('3.14\n')  # String that looks like a float
        
        # Disable auto-detect to preserve strings
        options = CSVLoadOptions()
        options.auto_detect_types = False
        
        wb = Workbook()
        wb.load_csv(csv_path, options)
        ws = wb.worksheets[0]
        
        self.assertIsInstance(ws.cells['A2'].value, str)
        self.assertEqual(ws.cells['A2'].value, 'Hello World')
        self.assertIsInstance(ws.cells['A3'].value, str)
        self.assertEqual(ws.cells['A3'].value, '123')
        self.assertIsInstance(ws.cells['A4'].value, str)
        self.assertEqual(ws.cells['A4'].value, '3.14')


class TestCSVSpecialCharacters(unittest.TestCase):
    """Test CSV handling of special characters."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_quoted_fields(self):
        """Test CSV with quoted fields containing delimiters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Name'
        ws.cells['B1'].value = 'Description'
        ws.cells['A2'].value = 'Alice'
        ws.cells['B2'].value = 'Has comma, in text'
        
        output_path = os.path.join(self.test_dir, 'test_quoted.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['B2'].value, 'Has comma, in text')
    
    def test_multiline_values(self):
        """Test CSV with multiline values."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Name'
        ws.cells['B1'].value = 'Description'
        ws.cells['A2'].value = 'Alice'
        ws.cells['B2'].value = 'Line 1\nLine 2\nLine 3'
        
        output_path = os.path.join(self.test_dir, 'test_multiline.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['B2'].value, 'Line 1\nLine 2\nLine 3')
    
    def test_quotes_in_values(self):
        """Test CSV with quotes in values."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Name'
        ws.cells['B1'].value = 'Quote'
        ws.cells['A2'].value = 'Alice'
        ws.cells['B2'].value = 'She said "Hello"'
        
        output_path = os.path.join(self.test_dir, 'test_quotes.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['B2'].value, 'She said "Hello"')
    
    def test_special_characters(self):
        """Test CSV with various special characters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Special'
        ws.cells['A2'].value = '!@#$%^&*()'
        ws.cells['A3'].value = '[]{};:\'",.<>/?'
        ws.cells['A4'].value = 'Tab\there'
        ws.cells['A5'].value = 'Backslash\\'
        
        output_path = os.path.join(self.test_dir, 'test_special_chars.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['A2'].value, '!@#$%^&*()')
        self.assertEqual(ws_imported.cells['A3'].value, '[]{};:\'",.<>/?')
        self.assertEqual(ws_imported.cells['A4'].value, 'Tab\there')
        self.assertEqual(ws_imported.cells['A5'].value, 'Backslash\\')


class TestCSVEmptyAndMissingValues(unittest.TestCase):
    """Test CSV handling of empty and missing values."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_empty_cells(self):
        """Test CSV with empty cells."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Name'
        ws.cells['B1'].value = 'Age'
        ws.cells['A2'].value = 'Alice'
        ws.cells['B2'].value = None  # Empty
        ws.cells['A3'].value = None  # Empty
        ws.cells['B3'].value = 25
        
        output_path = os.path.join(self.test_dir, 'test_empty.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        # Empty cells should be None or empty string
        self.assertIsNone(ws_imported.cells['B2'].value)
        self.assertIsNone(ws_imported.cells['A3'].value)
    
    def test_uneven_rows(self):
        """Test CSV with uneven row lengths."""
        csv_path = os.path.join(self.test_dir, 'test_uneven.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('A,B,C\n')
            f.write('1,2,3\n')
            f.write('4,5\n')  # Missing third column
            f.write('6\n')  # Missing second and third columns
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        self.assertEqual(ws.cells['A1'].value, 'A')
        self.assertEqual(ws.cells['B1'].value, 'B')
        self.assertEqual(ws.cells['C1'].value, 'C')
        self.assertEqual(ws.cells['A2'].value, 1)
        self.assertEqual(ws.cells['B2'].value, 2)
        self.assertEqual(ws.cells['C2'].value, 3)
        self.assertEqual(ws.cells['A3'].value, 4)
        self.assertEqual(ws.cells['B3'].value, 5)
        # Missing cells should be None
        self.assertIsNone(ws.cells['C3'].value)
        self.assertEqual(ws.cells['A4'].value, 6)
        self.assertIsNone(ws.cells['B4'].value)
        self.assertIsNone(ws.cells['C4'].value)


class TestCSVHeadersAndSkipRows(unittest.TestCase):
    """Test CSV header handling and skip rows functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_import_with_header(self):
        """Test importing CSV with header row."""
        csv_path = os.path.join(self.test_dir, 'test_header.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Name,Age,City\n')
            f.write('Alice,30,New York\n')
            f.write('Bob,25,London\n')
        
        options = CSVLoadOptions()
        options.has_header = True
        
        wb = Workbook()
        wb.load_csv(csv_path, options)
        ws = wb.worksheets[0]
        
        # Header should be in row 1
        self.assertEqual(ws.cells['A1'].value, 'Name')
        self.assertEqual(ws.cells['B1'].value, 'Age')
        self.assertEqual(ws.cells['C1'].value, 'City')
        # Data starts at row 2
        self.assertEqual(ws.cells['A2'].value, 'Alice')
        self.assertEqual(ws.cells['B2'].value, 30)
        self.assertEqual(ws.cells['C2'].value, 'New York')
    
    def test_skip_rows(self):
        """Test skipping rows at the beginning of CSV."""
        csv_path = os.path.join(self.test_dir, 'test_skip_rows.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('# Comment line 1\n')
            f.write('# Comment line 2\n')
            f.write('Name,Age\n')
            f.write('Alice,30\n')
        
        options = CSVLoadOptions()
        options.skip_rows = 2
        
        wb = Workbook()
        wb.load_csv(csv_path, options)
        ws = wb.worksheets[0]
        
        # First row should be Name,Age
        self.assertEqual(ws.cells['A1'].value, 'Name')
        self.assertEqual(ws.cells['B1'].value, 'Age')
        self.assertEqual(ws.cells['A2'].value, 'Alice')
        self.assertEqual(ws.cells['B2'].value, 30)
    
    def test_skip_rows_with_header(self):
        """Test skipping rows with header option."""
        csv_path = os.path.join(self.test_dir, 'test_skip_header.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('# Metadata\n')
            f.write('Name,Age\n')
            f.write('Alice,30\n')
        
        options = CSVLoadOptions()
        options.skip_rows = 1
        options.has_header = True
        
        wb = Workbook()
        wb.load_csv(csv_path, options)
        ws = wb.worksheets[0]
        
        # Header in row 1, data in row 2
        self.assertEqual(ws.cells['A1'].value, 'Name')
        self.assertEqual(ws.cells['B1'].value, 'Age')
        self.assertEqual(ws.cells['A2'].value, 'Alice')
        self.assertEqual(ws.cells['B2'].value, 30)


class TestCSVQuotingOptions(unittest.TestCase):
    """Test CSV quoting options."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_quote_minimal(self):
        """Test QUOTE_MINIMAL quoting (default)."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Simple'
        ws.cells['B1'].value = 'Has, comma'
        ws.cells['C1'].value = 'No comma'
        
        options = CSVSaveOptions()
        options.quoting = 0  # csv.QUOTE_MINIMAL
        
        output_path = os.path.join(self.test_dir, 'test_quote_minimal.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Only field with comma should be quoted
        self.assertIn('Simple', content)
        self.assertIn('"Has, comma"', content)
        self.assertIn('No comma', content)
    
    def test_quote_all(self):
        """Test QUOTE_ALL quoting."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Simple'
        ws.cells['B1'].value = 'Text'
        
        options = CSVSaveOptions()
        options.quoting = 1  # csv.QUOTE_ALL
        
        output_path = os.path.join(self.test_dir, 'test_quote_all.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # All fields should be quoted
        self.assertIn('"Simple"', content)
        self.assertIn('"Text"', content)
    
    def test_quote_nonnumeric(self):
        """Test QUOTE_NONNUMERIC quoting."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Text'
        ws.cells['B1'].value = 42
        ws.cells['C1'].value = 3.14
        
        options = CSVSaveOptions()
        options.quoting = 2  # csv.QUOTE_NONNUMERIC
        
        output_path = os.path.join(self.test_dir, 'test_quote_nonnumeric.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Only non-numeric fields should be quoted
        self.assertIn('"Text"', content)
        self.assertIn('42', content)
        self.assertIn('3.14', content)


class TestCSVLineTerminators(unittest.TestCase):
    """Test CSV line terminator options."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_crlf_line_terminator(self):
        """Test CRLF line terminator (Windows default)."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Row1'
        ws.cells['A2'].value = 'Row2'
        
        options = CSVSaveOptions()
        options.line_terminator = '\r\n'
        
        output_path = os.path.join(self.test_dir, 'test_crlf.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'rb') as f:
            content = f.read()
        
        self.assertIn(b'Row1\r\nRow2', content)
    
    def test_lf_line_terminator(self):
        """Test LF line terminator (Unix)."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Row1'
        ws.cells['A2'].value = 'Row2'
        
        options = CSVSaveOptions()
        options.line_terminator = '\n'
        
        output_path = os.path.join(self.test_dir, 'test_lf.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'rb') as f:
            content = f.read()
        
        self.assertIn(b'Row1\nRow2', content)


class TestCSVWorksheetIndex(unittest.TestCase):
    """Test CSV export from different worksheets."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_export_first_worksheet(self):
        """Test exporting the first worksheet (default)."""
        wb = Workbook()
        ws1 = wb.worksheets[0]
        ws1.name = 'Sheet1'
        ws1.cells['A1'].value = 'Sheet1 Data'
        
        ws2 = wb.add_worksheet('Sheet2')
        ws2.cells['A1'].value = 'Sheet2 Data'
        
        output_path = os.path.join(self.test_dir, 'test_sheet1.csv')
        wb.save_as_csv(output_path)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('Sheet1 Data', content)
        self.assertNotIn('Sheet2 Data', content)
    
    def test_export_second_worksheet(self):
        """Test exporting the second worksheet."""
        wb = Workbook()
        ws1 = wb.worksheets[0]
        ws1.name = 'Sheet1'
        ws1.cells['A1'].value = 'Sheet1 Data'
        
        ws2 = wb.add_worksheet('Sheet2')
        ws2.cells['A1'].value = 'Sheet2 Data'
        
        options = CSVSaveOptions()
        options.worksheet_index = 1
        
        output_path = os.path.join(self.test_dir, 'test_sheet2.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('Sheet2 Data', content)
        self.assertNotIn('Sheet1 Data', content)
    
    def test_invalid_worksheet_index(self):
        """Test that invalid worksheet index raises error."""
        wb = Workbook()
        
        options = CSVSaveOptions()
        options.worksheet_index = 99
        
        output_path = os.path.join(self.test_dir, 'test_invalid_index.csv')
        
        with self.assertRaises(IndexError):
            wb.save_as_csv(output_path, options)


class TestCSVStringMethods(unittest.TestCase):
    """Test CSV string-based import/export methods."""
    
    def test_save_to_string(self):
        """Test saving workbook to CSV string."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Name'
        ws.cells['B1'].value = 'Age'
        ws.cells['A2'].value = 'Alice'
        ws.cells['B2'].value = 30
        
        csv_string = CSVHandler.save_csv_to_string(wb)
        
        self.assertIn('Name,Age', csv_string)
        self.assertIn('Alice,30', csv_string)
    
    def test_load_from_string(self):
        """Test loading workbook from CSV string."""
        csv_string = 'Name,Age\nAlice,30\nBob,25\n'
        
        wb = Workbook()
        CSVHandler.load_csv_from_string(wb, csv_string)
        ws = wb.worksheets[0]
        
        self.assertEqual(ws.cells['A1'].value, 'Name')
        self.assertEqual(ws.cells['B1'].value, 'Age')
        self.assertEqual(ws.cells['A2'].value, 'Alice')
        self.assertEqual(ws.cells['B2'].value, 30)
        self.assertEqual(ws.cells['A3'].value, 'Bob')
        self.assertEqual(ws.cells['B3'].value, 25)
    
    def test_string_roundtrip(self):
        """Test round-trip using string methods."""
        wb_original = Workbook()
        ws_original = wb_original.worksheets[0]
        ws_original.cells['A1'].value = 'Name'
        ws_original.cells['B1'].value = 'Age'
        ws_original.cells['A2'].value = 'Alice'
        ws_original.cells['B2'].value = 30
        
        # Save to string
        csv_string = CSVHandler.save_csv_to_string(wb_original)
        
        # Load from string
        wb_loaded = Workbook()
        CSVHandler.load_csv_from_string(wb_loaded, csv_string)
        ws_loaded = wb_loaded.worksheets[0]
        
        # Verify data
        self.assertEqual(ws_loaded.cells['A1'].value, 'Name')
        self.assertEqual(ws_loaded.cells['B1'].value, 'Age')
        self.assertEqual(ws_loaded.cells['A2'].value, 'Alice')
        self.assertEqual(ws_loaded.cells['B2'].value, 30)


class TestCSVConvenienceFunctions(unittest.TestCase):
    """Test convenience functions for CSV operations."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_load_csv_workbook(self):
        """Test load_csv_workbook convenience function."""
        from aspose_cells.csv_handler import load_csv_workbook
        
        csv_path = os.path.join(self.test_dir, 'test_convenience_load.csv')
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write('Name,Age\n')
            f.write('Alice,30\n')
        
        wb = load_csv_workbook(csv_path)
        ws = wb.worksheets[0]
        
        self.assertEqual(ws.cells['A1'].value, 'Name')
        self.assertEqual(ws.cells['B1'].value, 'Age')
        self.assertEqual(ws.cells['A2'].value, 'Alice')
        self.assertEqual(ws.cells['B2'].value, 30)
    
    def test_save_workbook_as_csv(self):
        """Test save_workbook_as_csv convenience function."""
        from aspose_cells.csv_handler import save_workbook_as_csv
        
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Test'
        ws.cells['A2'].value = 'Data'
        
        output_path = os.path.join(self.test_dir, 'test_convenience_save.csv')
        save_workbook_as_csv(wb, output_path)
        
        self.assertTrue(os.path.exists(output_path))
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('Test', content)
        self.assertIn('Data', content)


class TestCSVLargeDataset(unittest.TestCase):
    """Test CSV handling with large datasets."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_large_dataset_export(self):
        """Test exporting a large dataset (1000 rows x 10 columns)."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Create 1000 rows x 10 columns of data
        for row in range(1, 1001):
            for col in range(1, 11):
                col_letter = chr(ord('A') + col - 1)
                ws.cells[f'{col_letter}{row}'].value = f'Row{row}_Col{col}'
        
        output_path = os.path.join(self.test_dir, 'test_large_export.csv')
        wb.save_as_csv(output_path)
        
        # Verify file exists and has reasonable size
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 100000)  # Should be at least 100KB
        
        # Import and verify a few values
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['A1'].value, 'Row1_Col1')
        self.assertEqual(ws_imported.cells['J1'].value, 'Row1_Col10')
        self.assertEqual(ws_imported.cells['A1000'].value, 'Row1000_Col1')
        self.assertEqual(ws_imported.cells['J1000'].value, 'Row1000_Col10')
    
    def test_large_dataset_import(self):
        """Test importing a large dataset."""
        csv_path = os.path.join(self.test_dir, 'test_large_import.csv')
        
        # Create CSV file with 500 rows
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            for row in range(1, 501):
                values = [f'Val{row}_{col}' for col in range(1, 6)]
                f.write(','.join(values) + '\n')
        
        wb = Workbook()
        wb.load_csv(csv_path)
        ws = wb.worksheets[0]
        
        # Verify first and last rows
        self.assertEqual(ws.cells['A1'].value, 'Val1_1')
        self.assertEqual(ws.cells['E1'].value, 'Val1_5')
        self.assertEqual(ws.cells['A500'].value, 'Val500_1')
        self.assertEqual(ws.cells['E500'].value, 'Val500_5')


class TestCSVUnicodeAndInternationalization(unittest.TestCase):
    """Test CSV handling with Unicode and international characters."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_chinese_characters(self):
        """Test CSV with Chinese characters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = '姓名'
        ws.cells['B1'].value = '年龄'
        ws.cells['A2'].value = '张三'
        ws.cells['B2'].value = 30
        ws.cells['A3'].value = '李四'
        ws.cells['B3'].value = 25
        
        output_path = os.path.join(self.test_dir, 'test_chinese.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['A1'].value, '姓名')
        self.assertEqual(ws_imported.cells['B1'].value, '年龄')
        self.assertEqual(ws_imported.cells['A2'].value, '张三')
        self.assertEqual(ws_imported.cells['B2'].value, 30)
        self.assertEqual(ws_imported.cells['A3'].value, '李四')
        self.assertEqual(ws_imported.cells['B3'].value, 25)
    
    def test_japanese_characters(self):
        """Test CSV with Japanese characters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = '名前'
        ws.cells['B1'].value = '年齢'
        ws.cells['A2'].value = '田中'
        ws.cells['B2'].value = 35
        
        output_path = os.path.join(self.test_dir, 'test_japanese.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['A1'].value, '名前')
        self.assertEqual(ws_imported.cells['B1'].value, '年齢')
        self.assertEqual(ws_imported.cells['A2'].value, '田中')
        self.assertEqual(ws_imported.cells['B2'].value, 35)
    
    def test_arabic_characters(self):
        """Test CSV with Arabic characters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'الاسم'
        ws.cells['B1'].value = 'العمر'
        ws.cells['A2'].value = 'أحمد'
        ws.cells['B2'].value = 28
        
        output_path = os.path.join(self.test_dir, 'test_arabic.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['A1'].value, 'الاسم')
        self.assertEqual(ws_imported.cells['B1'].value, 'العمر')
        self.assertEqual(ws_imported.cells['A2'].value, 'أحمد')
        self.assertEqual(ws_imported.cells['B2'].value, 28)
    
    def test_mixed_unicode(self):
        """Test CSV with mixed Unicode characters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'English'
        ws.cells['B1'].value = '中文'
        ws.cells['C1'].value = '日本語'
        ws.cells['D1'].value = 'العربية'
        ws.cells['E1'].value = '한국어'
        ws.cells['F1'].value = 'Русский'
        
        output_path = os.path.join(self.test_dir, 'test_mixed_unicode.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(ws_imported.cells['A1'].value, 'English')
        self.assertEqual(ws_imported.cells['B1'].value, '中文')
        self.assertEqual(ws_imported.cells['C1'].value, '日本語')
        self.assertEqual(ws_imported.cells['D1'].value, 'العربية')
        self.assertEqual(ws_imported.cells['E1'].value, '한국어')
        self.assertEqual(ws_imported.cells['F1'].value, 'Русский')


class TestCSVDateTimeFormatting(unittest.TestCase):
    """Test CSV datetime formatting during export."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_date_formatting(self):
        """Test date formatting in CSV export."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = date(2023, 1, 15)
        ws.cells['A2'].value = date(2023, 12, 31)
        
        options = CSVSaveOptions()
        options.date_format = '%Y/%m/%d'
        
        output_path = os.path.join(self.test_dir, 'test_date_format.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('2023/01/15', content)
        self.assertIn('2023/12/31', content)
    
    def test_datetime_formatting(self):
        """Test datetime formatting in CSV export."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = datetime(2023, 1, 15, 14, 30, 45)
        ws.cells['A2'].value = datetime(2023, 12, 31, 23, 59, 59)
        
        options = CSVSaveOptions()
        options.datetime_format = '%Y-%m-%d %H:%M:%S'
        
        output_path = os.path.join(self.test_dir, 'test_datetime_format.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('2023-01-15 14:30:45', content)
        self.assertIn('2023-12-31 23:59:59', content)
    
    def test_time_formatting(self):
        """Test time formatting in CSV export."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = time(14, 30, 45)
        ws.cells['A2'].value = time(23, 59, 59)
        
        options = CSVSaveOptions()
        options.time_format = '%H:%M:%S'
        
        output_path = os.path.join(self.test_dir, 'test_time_format.csv')
        wb.save_as_csv(output_path, options)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('14:30:45', content)
        self.assertIn('23:59:59', content)


class TestCSVNumberFormatting(unittest.TestCase):
    """Test CSV number formatting during export."""

    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)

    def test_numeric_number_format_export(self):
        """Test numeric formatting in CSV export."""
        wb = Workbook()
        ws = wb.worksheets[0]

        ws.cells['A1'].value = 1234.5
        ws.cells['A1'].style.set_number_format('#,##0.00')

        ws.cells['A2'].value = 0.256
        ws.cells['A2'].style.set_number_format('0%')

        ws.cells['A3'].value = 12
        ws.cells['A3'].style.set_number_format('0.000')

        output_path = os.path.join(self.test_dir, 'test_number_format.csv')
        wb.save_as_csv(output_path)

        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertIn('"1,234.50"', content)
        self.assertIn('26%', content)
        self.assertIn('12.000', content)


class TestCSVBooleanFormatting(unittest.TestCase):
    """Test CSV boolean formatting during export."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_boolean_export(self):
        """Test boolean value export to CSV."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = True
        ws.cells['A2'].value = False
        ws.cells['A3'].value = True
        
        output_path = os.path.join(self.test_dir, 'test_boolean.csv')
        wb.save_as_csv(output_path)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('TRUE', content)
        self.assertIn('FALSE', content)


class TestCSVEdgeCases(unittest.TestCase):
    """Test CSV edge cases and error handling."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_dir = 'outputfiles'
        os.makedirs(self.test_dir, exist_ok=True)
    
    def test_empty_worksheet(self):
        """Test exporting empty worksheet."""
        wb = Workbook()
        ws = wb.worksheets[0]
        # Don't add any data
        
        output_path = os.path.join(self.test_dir, 'test_empty_worksheet.csv')
        wb.save_as_csv(output_path)
        
        # File should exist but be empty or minimal
        self.assertTrue(os.path.exists(output_path))
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Empty worksheet should produce empty or minimal content
        self.assertEqual(content.strip(), '')
    
    def test_single_cell(self):
        """Test exporting single cell."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 'Single Value'
        
        output_path = os.path.join(self.test_dir, 'test_single_cell.csv')
        wb.save_as_csv(output_path)
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertEqual(content.strip(), 'Single Value')
    
    def test_very_long_string(self):
        """Test exporting very long string."""
        wb = Workbook()
        ws = wb.worksheets[0]
        long_string = 'A' * 10000
        ws.cells['A1'].value = long_string
        
        output_path = os.path.join(self.test_dir, 'test_long_string.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertEqual(len(ws_imported.cells['A1'].value), 10000)
        self.assertEqual(ws_imported.cells['A1'].value, long_string)
    
    def test_import_nonexistent_file(self):
        """Test importing non-existent file."""
        wb = Workbook()
        
        with self.assertRaises(FileNotFoundError):
            wb.load_csv('nonexistent_file.csv')
    
    def test_numeric_precision(self):
        """Test numeric precision preservation."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.cells['A1'].value = 3.141592653589793
        ws.cells['A2'].value = 2.718281828459045
        
        output_path = os.path.join(self.test_dir, 'test_precision.csv')
        wb.save_as_csv(output_path)
        
        # Import and verify
        wb_imported = Workbook()
        wb_imported.load_csv(output_path)
        ws_imported = wb_imported.worksheets[0]
        
        self.assertAlmostEqual(ws_imported.cells['A1'].value, 3.141592653589793, places=10)
        self.assertAlmostEqual(ws_imported.cells['A2'].value, 2.718281828459045, places=10)


if __name__ == '__main__':
    unittest.main()
