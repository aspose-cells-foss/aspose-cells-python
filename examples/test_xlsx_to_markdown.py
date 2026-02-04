"""
Test Suite for XLSX to Markdown Conversion

This test suite covers converting XLSX files to Markdown format, including:
- Basic XLSX to Markdown conversion
- Multiple worksheets handling
- Custom formatting options
- Unicode and special characters
- Large datasets
- Empty worksheets
"""

import unittest
import os
import sys
from datetime import datetime, date, time

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook
from aspose_cells.markdown_handler import MarkdownHandler, MarkdownSaveOptions


class TestXLSXToMarkdownConversion(unittest.TestCase):
    """Test XLSX to Markdown conversion functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.input_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'input')
        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputfiles')
        os.makedirs(self.output_dir, exist_ok=True)
    
    def test_convert_all_xlsx_to_markdown(self):
        """
        Test converting all XLSX files from input directory to Markdown.
        
        This test:
        1. Finds all .xlsx files in the input directory
        2. Converts each to Markdown format
        3. Saves to tests/outputfiles directory
        4. Verifies the Markdown files were created
        """
        # Find all XLSX files in input directory
        xlsx_files = []
        if os.path.exists(self.input_dir):
            for filename in os.listdir(self.input_dir):
                if filename.endswith('.xlsx'):
                    xlsx_files.append(filename)
        
        self.assertGreater(len(xlsx_files), 0, "No XLSX files found in input directory")
        
        # Convert each XLSX file to Markdown
        for xlsx_file in xlsx_files:
            # Input file path
            input_path = os.path.join(self.input_dir, xlsx_file)
            
            # Output file path (change extension to .md)
            base_name = os.path.splitext(xlsx_file)[0]
            output_path = os.path.join(self.output_dir, f'{base_name}.md')
            
            # Load workbook
            wb = Workbook(input_path)
            
            # Save as Markdown with default options
            wb.save_as_markdown(output_path)
            
            # Verify file was created
            self.assertTrue(os.path.exists(output_path), 
                          f"Markdown file not created: {output_path}")
            
            # Verify file is not empty
            file_size = os.path.getsize(output_path)
            self.assertGreater(file_size, 0, 
                           f"Markdown file is empty: {output_path}")
            
            # Read and verify basic Markdown structure
            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Markdown should have some structure (headers, tables, etc.)
            self.assertGreater(len(content), 0, 
                           f"Markdown file has no content: {output_path}")
            
            print(f"[OK] Converted {xlsx_file} to {base_name}.md")


class TestMarkdownExportOptions(unittest.TestCase):
    """Test Markdown export with various options."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.output_dir = 'tests/outputfiles'
        os.makedirs(self.output_dir, exist_ok=True)
    
    def test_markdown_with_custom_alignment(self):
        """Test Markdown export with custom column alignment."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['C1'].value = "Active"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 30
        ws.cells['C2'].value = True
        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 25
        ws.cells['C3'].value = False
        
        # Create options with center alignment
        options = MarkdownSaveOptions()
        options.default_alignment = 'center'
        
        output_path = os.path.join(self.output_dir, 'test_markdown_alignment.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify alignment markers in content
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Center alignment should have :---: separators
        self.assertIn(':---:', content)
    
    def test_markdown_without_worksheet_name(self):
        """Test Markdown export without worksheet name header."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "TestData"
        
        # Add data
        ws.cells['A1'].value = "Name"
        ws.cells['A2'].value = "Alice"
        
        # Create options without worksheet name
        options = MarkdownSaveOptions()
        options.include_worksheet_name = False
        
        output_path = os.path.join(self.output_dir, 'test_markdown_no_header.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify no worksheet name header
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Should not have ## TestData header
        self.assertNotIn('## TestData', content)
    
    def test_markdown_with_custom_header_level(self):
        """Test Markdown export with custom header level."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "MySheet"
        
        # Add data
        ws.cells['A1'].value = "Data"
        ws.cells['A2'].value = "Value"
        
        # Create options with H3 header
        options = MarkdownSaveOptions()
        options.header_level = 3
        
        output_path = os.path.join(self.output_dir, 'test_markdown_h3.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify H3 header
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Should have ### MySheet (not ## MySheet)
        self.assertIn('### MySheet', content)
        self.assertNotIn('\n## MySheet', content)  # Check for H2 at line start
    
    def test_markdown_with_date_formatting(self):
        """Test Markdown export with custom date formatting."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data with dates
        ws.cells['A1'].value = "Date"
        ws.cells['A2'].value = date(2023, 1, 15)
        ws.cells['A3'].value = date(2023, 12, 31)
        
        # Create options with custom date format
        options = MarkdownSaveOptions()
        options.date_format = '%Y/%m/%d'
        
        output_path = os.path.join(self.output_dir, 'test_markdown_dates.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify date format
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Should have dates in YYYY/MM/DD format
        self.assertIn('2023/01/15', content)
        self.assertIn('2023/12/31', content)
    
    def test_markdown_with_datetime_formatting(self):
        """Test Markdown export with custom datetime formatting."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data with datetimes
        ws.cells['A1'].value = "DateTime"
        ws.cells['A2'].value = datetime(2023, 1, 15, 14, 30, 45)
        ws.cells['A3'].value = datetime(2023, 12, 31, 23, 59, 59)
        
        # Create options with custom datetime format
        options = MarkdownSaveOptions()
        options.datetime_format = '%Y-%m-%d %H:%M:%S'
        
        output_path = os.path.join(self.output_dir, 'test_markdown_datetimes.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify datetime format
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Should have datetimes in YYYY-MM-DD HH:MM:SS format
        self.assertIn('2023-01-15 14:30:45', content)
        self.assertIn('2023-12-31 23:59:59', content)
    
    def test_markdown_with_boolean_values(self):
        """Test Markdown export with boolean values."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data with booleans
        ws.cells['A1'].value = "Active"
        ws.cells['A2'].value = True
        ws.cells['A3'].value = False
        ws.cells['A4'].value = True
        
        output_path = os.path.join(self.output_dir, 'test_markdown_booleans.md')
        wb.save_as_markdown(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify boolean formatting
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Booleans should be Yes/No
        self.assertIn('Yes', content)
        self.assertIn('No', content)
    
    def test_markdown_with_empty_cells(self):
        """Test Markdown export with empty cells."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data with empty cells
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = None  # Empty
        ws.cells['A3'].value = None  # Empty
        ws.cells['B3'].value = 25
        
        output_path = os.path.join(self.output_dir, 'test_markdown_empty.md')
        wb.save_as_markdown(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify table structure
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Should have proper table structure
        self.assertIn('|', content)  # Table separators
        self.assertIn('Name', content)
        self.assertIn('Age', content)
    
    def test_markdown_with_special_characters(self):
        """Test Markdown export with special characters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data with special characters
        ws.cells['A1'].value = "Text"
        ws.cells['A2'].value = "Has | pipe"
        ws.cells['A3'].value = "Has * asterisk"
        ws.cells['A4'].value = "Has _ underscore"
        
        output_path = os.path.join(self.output_dir, 'test_markdown_special.md')
        wb.save_as_markdown(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify special character handling
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Pipes should be escaped by default
        self.assertIn('\\|', content)
    
    def test_markdown_with_unicode(self):
        """Test Markdown export with Unicode characters."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data with various Unicode
        ws.cells['A1'].value = "Language"
        ws.cells['A2'].value = "English"
        ws.cells['A3'].value = "中文"
        ws.cells['A4'].value = "日本語"
        ws.cells['A5'].value = "العربية"
        ws.cells['A6'].value = "한국어"
        
        output_path = os.path.join(self.output_dir, 'test_markdown_unicode.md')
        wb.save_as_markdown(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify Unicode characters are preserved
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('中文', content)
        self.assertIn('日本語', content)
        self.assertIn('العربية', content)
        self.assertIn('한국어', content)
    
    def test_markdown_multiple_worksheets(self):
        """Test Markdown export with multiple worksheets."""
        wb = Workbook()
        
        # First worksheet
        ws1 = wb.worksheets[0]
        ws1.name = "Sheet1"
        ws1.cells['A1'].value = "Data"
        ws1.cells['A2'].value = "Value1"
        
        # Second worksheet
        ws2 = wb.add_worksheet("Sheet2")
        ws2.cells['A1'].value = "Data"
        ws2.cells['A2'].value = "Value2"
        
        # Export all worksheets
        options = MarkdownSaveOptions()
        options.worksheet_index = -1  # Export all
        
        output_path = os.path.join(self.output_dir, 'test_markdown_multiple_sheets.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify both worksheets are included
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('## Sheet1', content)
        self.assertIn('## Sheet2', content)
        self.assertIn('Value1', content)
        self.assertIn('Value2', content)
    
    def test_markdown_single_worksheet(self):
        """Test Markdown export of single worksheet from multi-sheet workbook."""
        wb = Workbook()
        
        # First worksheet
        ws1 = wb.worksheets[0]
        ws1.name = "Sheet1"
        ws1.cells['A1'].value = "Data"
        ws1.cells['A2'].value = "Value1"
        
        # Second worksheet
        ws2 = wb.add_worksheet("Sheet2")
        ws2.cells['A1'].value = "Data"
        ws2.cells['A2'].value = "Value2"
        
        # Export only second worksheet
        options = MarkdownSaveOptions()
        options.worksheet_index = 1  # Export only Sheet2
        
        output_path = os.path.join(self.output_dir, 'test_markdown_single_sheet.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify only Sheet2 is included
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertNotIn('## Sheet1', content)
        self.assertIn('## Sheet2', content)
        self.assertNotIn('Value1', content)
        self.assertIn('Value2', content)
    
    def test_markdown_empty_worksheet(self):
        """Test Markdown export of empty worksheet."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "EmptySheet"
        # Don't add any data
        
        output_path = os.path.join(self.output_dir, 'test_markdown_empty_sheet.md')
        wb.save_as_markdown(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify content indicates no data
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('EmptySheet', content)
        self.assertIn('*No data*', content)
    
    def test_markdown_large_dataset(self):
        """Test Markdown export with large dataset."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Create large dataset (100 rows x 10 columns)
        for row in range(1, 101):
            for col in range(1, 11):
                col_letter = chr(ord('A') + col - 1)
                ws.cells[f'{col_letter}{row}'].value = f'Row{row}_Col{col}'
        
        output_path = os.path.join(self.output_dir, 'test_markdown_large.md')
        wb.save_as_markdown(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify file has reasonable size
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 1000)  # At least 1KB
        
        # Verify some data points
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self.assertIn('Row1_Col1', content)
        self.assertIn('Row100_Col10', content)
    
    def test_markdown_to_string(self):
        """Test Markdown export to string."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data
        ws.cells['A1'].value = "Name"
        ws.cells['A2'].value = "Alice"
        
        # Export to string
        md_string = MarkdownHandler.save_markdown_to_string(wb)
        
        # Verify string content
        self.assertGreater(len(md_string), 0)
        self.assertIn('Name', md_string)
        self.assertIn('Alice', md_string)
        self.assertIn('|', md_string)  # Table separators
    
    def test_markdown_with_row_numbers(self):
        """Test Markdown export with row numbers."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data
        ws.cells['A1'].value = "Name"
        ws.cells['A2'].value = "Alice"
        ws.cells['A3'].value = "Bob"
        
        # Create options with row numbers
        options = MarkdownSaveOptions()
        options.include_row_numbers = True
        
        output_path = os.path.join(self.output_dir, 'test_markdown_row_numbers.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify row numbers
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Should have row number column
        # Note: Header row (A1) becomes table header, data rows (A2, A3) are numbered 1 and 2
        self.assertIn('#', content)  # Row number header
        self.assertIn('1', content)
        self.assertIn('2', content)
    
    def test_markdown_with_max_column_width(self):
        """Test Markdown export with maximum column width."""
        wb = Workbook()
        ws = wb.worksheets[0]
        
        # Add data with long strings
        ws.cells['A1'].value = "Short"
        ws.cells['A2'].value = "This is a very long string that should be truncated"
        
        # Create options with max width
        options = MarkdownSaveOptions()
        options.max_column_width = 20
        
        output_path = os.path.join(self.output_dir, 'test_markdown_max_width.md')
        wb.save_as_markdown(output_path, options)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify truncation
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Long string should be truncated with '...'
        self.assertIn('...', content)
        # Full string should not be present
        self.assertNotIn('This is a very long string that should be truncated', content)


if __name__ == '__main__':
    unittest.main()
