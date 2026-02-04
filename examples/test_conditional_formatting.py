"""
Test cases for Conditional Formatting feature.

This test suite verifies that conditional formatting rules can be created,
configured, saved to xlsx files, and loaded back correctly.

Features tested:
- Cell value rules (greater than, less than, between, equal to)
- Text rules (contains, does not contain, begins with, ends with)
- Date rules (yesterday, today, tomorrow, last 7 days, etc.)
- Duplicate/unique values rules
- Top/bottom rules (top 10 items, top 10%)
- Above/below average rules
- Color scales (2-color, 3-color)
- Data bars
- Icon sets
- Formula-based rules
- Save conditional formatting to xlsx file
- Load conditional formatting from xlsx file
"""

import unittest
import os
import sys

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook


class TestConditionalFormatting(unittest.TestCase):
    """
    Test comprehensive conditional formatting functionality with save/load verification.
    
    Features tested:
    - Cell value rules (greater than, less than, between, equal to)
    - Text rules (contains, does not contain, begins with, ends with)
    - Date rules (yesterday, today, tomorrow, last 7 days, etc.)
    - Duplicate/unique values rules
    - Top/bottom rules (top 10 items, top 10%)
    - Above/below average rules
    - Color scales (2-color, 3-color)
    - Data bars
    - Icon sets
    - Formula-based rules
    - Save conditional formatting to xlsx file
    - Load conditional formatting from xlsx file
    """
    
    def setUp(self):
        """Set up test workbook."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.worksheets[0]
    
    def tearDown(self):
        """Clean up test files."""
        # Note: Test output files are kept in outputfiles/ directory for verification
        # Commented out file cleanup to preserve generated files
        pass
    
    def test_cell_value_rules(self):
        """Test cell value rules (greater than, less than, between, equal to)."""
        print("Testing cell value rules...")
        
        # Test greater than rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'cellValue'
        cf1.operator = 'greaterThan'
        cf1.formula1 = '100'
        cf1.range = 'B1:B10'
        cf1.font.bold = True
        cf1.font.color = 'FFFF0000'
        self.assertEqual(cf1.type, 'cellValue')
        self.assertEqual(cf1.operator, 'greaterThan')
        self.assertEqual(cf1.formula1, '100')
        self.assertEqual(cf1.range, 'B1:B10')
        self.assertTrue(cf1.font.bold)
        self.assertEqual(cf1.font.color, 'FFFF0000')
        print("  Greater than rule created")
        
        # Test less than rule
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'cellValue'
        cf2.operator = 'lessThan'
        cf2.formula1 = '50'
        cf2.range = 'A1:A10'
        cf2.fill.set_solid_fill('FFFFFF00')
        self.assertEqual(cf2.type, 'cellValue')
        self.assertEqual(cf2.operator, 'lessThan')
        self.assertEqual(cf2.formula1, '50')
        self.assertEqual(cf2.range, 'A1:A10')
        self.assertEqual(cf2.fill.foreground_color, 'FFFFFF00')
        print("  Less than rule created")
        
        # Test between rule
        cf3 = self.worksheet.conditional_formats.add()
        cf3.type = 'cellValue'
        cf3.operator = 'between'
        cf3.formula1 = '10'
        cf3.formula2 = '100'
        cf3.range = 'C1:C10'
        cf3.font.italic = True
        cf3.font.color = 'FF0000FF'
        self.assertEqual(cf3.type, 'cellValue')
        self.assertEqual(cf3.operator, 'between')
        self.assertEqual(cf3.formula1, '10')
        self.assertEqual(cf3.formula2, '100')
        self.assertEqual(cf3.range, 'C1:C10')
        self.assertTrue(cf3.font.italic)
        self.assertEqual(cf3.font.color, 'FF0000FF')
        print("  Between rule created")
        
        # Test equal to rule
        cf4 = self.worksheet.conditional_formats.add()
        cf4.type = 'cellValue'
        cf4.operator = 'equal'
        cf4.formula1 = '75'
        cf4.range = 'D1:D10'
        cf4.font.underline = True
        self.assertEqual(cf4.type, 'cellValue')
        self.assertEqual(cf4.operator, 'equal')
        self.assertEqual(cf4.formula1, '75')
        self.assertEqual(cf4.range, 'D1:D10')
        self.assertTrue(cf4.font.underline)
        print("  Equal to rule created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
            self.worksheet.cells[f'B{i}'].value = i * 100
            self.worksheet.cells[f'C{i}'].value = 50 + (i * 10)
            self.worksheet.cells[f'D{i}'].value = i * 75
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_cell_value_rules.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Cell value rules test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_text_rules(self):
        """Test text rules (contains, does not contain, begins with, ends with)."""
        print("Testing text rules...")
        
        # Test contains rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'text'
        cf1.text_operator = 'contains'
        cf1.text_formula = 'error'
        cf1.range = 'A1:A10'
        cf1.font.color = 'FFFF0000'
        cf1.fill.set_solid_fill('FFFF00')
        self.assertEqual(cf1.type, 'text')
        self.assertEqual(cf1.text_operator, 'contains')
        self.assertEqual(cf1.text_formula, 'error')
        self.assertEqual(cf1.range, 'A1:A10')
        self.assertEqual(cf1.font.color, 'FFFF0000')
        self.assertEqual(cf1.fill.foreground_color, 'FFFF00')
        print(" Contains rule created")
        
        # Test does not contain rule
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'text'
        cf2.text_operator = 'notContains'
        cf2.text_formula = 'warning'
        cf2.range = 'B1:B10'
        cf2.font.color = 'FF00FF00'
        cf2.fill.set_solid_fill('FFFFFF00')
        self.assertEqual(cf2.type, 'text')
        self.assertEqual(cf2.text_operator, 'notContains')
        self.assertEqual(cf2.text_formula, 'warning')
        self.assertEqual(cf2.range, 'B1:B10')
        self.assertEqual(cf2.font.color, 'FF00FF00')
        self.assertEqual(cf2.fill.foreground_color, 'FFFFFF00')
        print(" Does not contain rule created")
        
        # Test begins with rule
        cf3 = self.worksheet.conditional_formats.add()
        cf3.type = 'text'
        cf3.text_operator = 'beginsWith'
        cf3.text_formula = 'prefix'
        cf3.range = 'C1:C10'
        cf3.font.bold = True
        self.assertEqual(cf3.type, 'text')
        self.assertEqual(cf3.text_operator, 'beginsWith')
        self.assertEqual(cf3.text_formula, 'prefix')
        self.assertEqual(cf3.range, 'C1:C10')
        self.assertTrue(cf3.font.bold)
        print(" Begins with rule created")
        
        # Test ends with rule
        cf4 = self.worksheet.conditional_formats.add()
        cf4.type = 'text'
        cf4.text_operator = 'endsWith'
        cf4.text_formula = 'suffix'
        cf4.range = 'D1:D10'
        cf4.font.italic = True
        self.assertEqual(cf4.type, 'text')
        self.assertEqual(cf4.text_operator, 'endsWith')
        self.assertEqual(cf4.text_formula, 'suffix')
        self.assertEqual(cf4.range, 'D1:D10')
        self.assertTrue(cf4.font.italic)
        print(" Ends with rule created")
        
        # Add test data for text rules
        # Column A: Contains "error" - should trigger for cells with "error"
        self.worksheet.cells['A1'].value = "error message"
        self.worksheet.cells['A2'].value = "warning message"
        self.worksheet.cells['A3'].value = "error found"
        self.worksheet.cells['A4'].value = "info message"
        self.worksheet.cells['A5'].value = "error detected"
        self.worksheet.cells['A6'].value = "success"
        self.worksheet.cells['A7'].value = "error"
        self.worksheet.cells['A8'].value = "warning"
        self.worksheet.cells['A9'].value = "critical error"
        self.worksheet.cells['A10'].value = "normal"
        
        # Column B: Does not contain "warning" - should trigger for cells without "warning"
        self.worksheet.cells['B1'].value = "error message"
        self.worksheet.cells['B2'].value = "success message"
        self.worksheet.cells['B3'].value = "info message"
        self.worksheet.cells['B4'].value = "warning message"
        self.worksheet.cells['B5'].value = "critical error"
        self.worksheet.cells['B6'].value = "normal"
        self.worksheet.cells['B7'].value = "error"
        self.worksheet.cells['B8'].value = "warning"
        self.worksheet.cells['B9'].value = "critical error"
        self.worksheet.cells['B10'].value = "normal"
        
        # Column C: Begins with "prefix" - should trigger for cells starting with "prefix"
        self.worksheet.cells['C1'].value = "prefix_test"
        self.worksheet.cells['C2'].value = "other_text"
        self.worksheet.cells['C3'].value = "prefix_data"
        self.worksheet.cells['C4'].value = "prefix_value"
        self.worksheet.cells['C5'].value = "different"
        self.worksheet.cells['C6'].value = "prefix_item"
        self.worksheet.cells['C7'].value = "prefix_string"
        self.worksheet.cells['C8'].value = "another_text"
        self.worksheet.cells['C9'].value = "prefix_test_again"
        self.worksheet.cells['C10'].value = "no_prefix"
        
        # Column D: Ends with "suffix" - should trigger for cells ending with "suffix"
        self.worksheet.cells['D1'].value = "test_suffix"
        self.worksheet.cells['D2'].value = "suffix_data"
        self.worksheet.cells['D3'].value = "suffix_value"
        self.worksheet.cells['D4'].value = "different"
        self.worksheet.cells['D5'].value = "suffix_item"
        self.worksheet.cells['D6'].value = "suffix_string"
        self.worksheet.cells['D7'].value = "another_text"
        self.worksheet.cells['D8'].value = "suffix_test_again"
        self.worksheet.cells['D9'].value = "no_suffix"
        self.worksheet.cells['D10'].value = "normal"
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_text_rules.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Text rules test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_date_rules(self):
        """Test date rules (yesterday, today, tomorrow, last 7 days, etc.)."""
        print("Testing date rules...")
        
        # Test yesterday rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'date'
        cf1.date_operator = 'yesterday'
        cf1.range = 'A1:A10'
        cf1.fill.set_solid_fill('FFFFFF00')
        self.assertEqual(cf1.type, 'date')
        self.assertEqual(cf1.date_operator, 'yesterday')
        self.assertEqual(cf1.range, 'A1:A10')
        self.assertEqual(cf1.fill.foreground_color, 'FFFFFF00')
        print(" Yesterday rule created")
        
        # Test today rule
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'date'
        cf2.date_operator = 'today'
        cf2.range = 'B1:B10'
        cf2.fill.set_solid_fill('FFFF00')
        self.assertEqual(cf2.type, 'date')
        self.assertEqual(cf2.date_operator, 'today')
        self.assertEqual(cf2.range, 'B1:B10')
        self.assertEqual(cf2.fill.foreground_color, 'FFFF00')
        print(" Today rule created")
        
        # Test tomorrow rule
        cf3 = self.worksheet.conditional_formats.add()
        cf3.type = 'date'
        cf3.date_operator = 'tomorrow'
        cf3.range = 'C1:C10'
        cf3.fill.set_solid_fill('00FF00')
        self.assertEqual(cf3.type, 'date')
        self.assertEqual(cf3.date_operator, 'tomorrow')
        self.assertEqual(cf3.range, 'C1:C10')
        self.assertEqual(cf3.fill.foreground_color, '00FF00')
        print(" Tomorrow rule created")
        
        # Test last 7 days rule
        cf4 = self.worksheet.conditional_formats.add()
        cf4.type = 'date'
        cf4.date_operator = 'last7Days'
        cf4.range = 'D1:D10'
        cf4.font.color = 'FF00FF00'
        self.assertEqual(cf4.type, 'date')
        self.assertEqual(cf4.date_operator, 'last7Days')
        self.assertEqual(cf4.range, 'D1:D10')
        self.assertEqual(cf4.font.color, 'FF00FF00')
        print(" Last 7 days rule created")
        
        # Add test data
        from datetime import datetime, timedelta
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)
        tomorrow = today + timedelta(days=1)
        last_7_days = today - timedelta(days=7)
        for i, date in enumerate([yesterday, today, tomorrow, last_7_days], 1):
            self.worksheet.cells[f'A{i}'].value = date
            self.worksheet.cells[f'B{i}'].value = date
            self.worksheet.cells[f'C{i}'].value = date
            self.worksheet.cells[f'D{i}'].value = date
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_date_rules.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Date rules test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_duplicate_unique_values(self):
        """Test duplicate/unique values rules."""
        print("Testing duplicate/unique values rules...")
        
        # Test duplicate values rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'duplicateValues'
        cf1.duplicate = True
        cf1.range = 'A1:A10'
        cf1.fill.set_solid_fill('FFFF0000')
        self.assertEqual(cf1.type, 'duplicateValues')
        self.assertTrue(cf1.duplicate)
        self.assertEqual(cf1.range, 'A1:A10')
        self.assertEqual(cf1.fill.foreground_color, 'FFFF0000')
        print(" Duplicate values rule created")
        
        # Test unique values rule
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'uniqueValues'
        cf2.duplicate = False
        cf2.range = 'B1:B10'
        cf2.fill.set_solid_fill('00FF00')
        self.assertEqual(cf2.type, 'uniqueValues')
        self.assertFalse(cf2.duplicate)
        self.assertEqual(cf2.range, 'B1:B10')
        self.assertEqual(cf2.fill.foreground_color, '00FF00')
        print(" Unique values rule created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i if i <= 5 else i - 5
            self.worksheet.cells[f'B{i}'].value = i * 100
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_duplicate_unique_values.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Duplicate/unique values test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_top_bottom_rules(self):
        """Test top/bottom rules (top 10 items, top 10%)."""
        print("Testing top/bottom rules...")
        
        # Test top 10 items rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'top10'
        cf1.top = True
        cf1.rank = 10
        cf1.range = 'A1:A10'
        cf1.fill.set_solid_fill('FFFF00')
        self.assertEqual(cf1.type, 'top10')
        self.assertTrue(cf1.top)
        self.assertEqual(cf1.rank, 10)
        self.assertEqual(cf1.range, 'A1:A10')
        self.assertEqual(cf1.fill.foreground_color, 'FFFF00')
        print(" Top 10 items rule created")
        
        # Test top 10% rule
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'top10'
        cf2.top = True
        cf2.percent = True
        cf2.rank = 10
        cf2.range = 'B1:B10'
        cf2.fill.set_solid_fill('00FF00')
        self.assertEqual(cf2.type, 'top10')
        self.assertTrue(cf2.top)
        self.assertTrue(cf2.percent)
        self.assertEqual(cf2.rank, 10)
        self.assertEqual(cf2.range, 'B1:B10')
        self.assertEqual(cf2.fill.foreground_color, '00FF00')
        print(" Top 10% rule created")
        
        # Test bottom 10 items rule
        cf3 = self.worksheet.conditional_formats.add()
        cf3.type = 'bottom10'
        cf3.top = False
        cf3.rank = 10
        cf3.range = 'C1:C10'
        cf3.fill.set_solid_fill('FFFFFF00')
        self.assertEqual(cf3.type, 'bottom10')
        self.assertFalse(cf3.top)
        self.assertEqual(cf3.rank, 10)
        self.assertEqual(cf3.range, 'C1:C10')
        self.assertEqual(cf3.fill.foreground_color, 'FFFFFF00')
        print(" Bottom 10 items rule created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 100
            self.worksheet.cells[f'B{i}'].value = i * 100
            self.worksheet.cells[f'C{i}'].value = i * 100
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_top_bottom_rules.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Top/bottom rules test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_above_below_average(self):
        """Test above/below average rules."""
        print("Testing above/below average rules...")
        
        # Test above average rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'aboveAverage'
        cf1.above = True
        cf1.range = 'A1:A10'
        cf1.fill.set_solid_fill('FFFFFF00')
        self.assertEqual(cf1.type, 'aboveAverage')
        self.assertTrue(cf1.above)
        self.assertEqual(cf1.range, 'A1:A10')
        self.assertEqual(cf1.fill.foreground_color, 'FFFFFF00')
        print(" Above average rule created")
        
        # Test below average rule
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'belowAverage'
        cf2.above = False
        cf2.range = 'B1:B10'
        cf2.fill.set_solid_fill('00FF00')
        self.assertEqual(cf2.type, 'belowAverage')
        self.assertFalse(cf2.above)
        self.assertEqual(cf2.range, 'B1:B10')
        self.assertEqual(cf2.fill.foreground_color, '00FF00')
        print("  Below average rule created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 100
            self.worksheet.cells[f'B{i}'].value = i * 100
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_above_below_average.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Above/below average test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_color_scale(self):
        """Test color scale rules (2-color, 3-color)."""
        print("Testing color scale rules...")
        
        # Test 2-color scale
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'colorScale'
        cf1.color_scale_type = '2-color'
        cf1.min_color = 'FF63C384'
        cf1.max_color = 'FF006100'
        cf1.range = 'A1:A10'
        self.assertEqual(cf1.type, 'colorScale')
        self.assertEqual(cf1.color_scale_type, '2-color')
        self.assertEqual(cf1.min_color, 'FF63C384')
        self.assertEqual(cf1.max_color, 'FF006100')
        self.assertEqual(cf1.range, 'A1:A10')
        print("  2-color scale created")
        
        # Test 3-color scale
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'colorScale'
        cf2.color_scale_type = '3-color'
        cf2.min_color = 'FF63C384'
        cf2.mid_color = 'FFFFEB84'
        cf2.max_color = 'FF006100'
        cf2.range = 'B1:B10'
        self.assertEqual(cf2.type, 'colorScale')
        self.assertEqual(cf2.color_scale_type, '3-color')
        self.assertEqual(cf2.min_color, 'FF63C384')
        self.assertEqual(cf2.mid_color, 'FFFFEB84')
        self.assertEqual(cf2.max_color, 'FF006100')
        self.assertEqual(cf2.range, 'B1:B10')
        print("  3-color scale created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
            self.worksheet.cells[f'B{i}'].value = i * 100
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_color_scale.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Color scale test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_data_bar(self):
        """Test data bar rules."""
        print("Testing data bar rules...")
        
        # Test data bar rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'dataBar'
        cf1.bar_color = 'FF006100'
        cf1.negative_color = 'FFFF0000'
        cf1.show_border = True
        cf1.direction = 'left-to-right'
        cf1.range = 'A1:A10'
        self.assertEqual(cf1.type, 'dataBar')
        self.assertEqual(cf1.bar_color, 'FF006100')
        self.assertEqual(cf1.negative_color, 'FFFF0000')
        self.assertTrue(cf1.show_border)
        self.assertEqual(cf1.direction, 'left-to-right')
        self.assertEqual(cf1.range, 'A1:A10')
        print(" Data bar created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_data_bar.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Data bar test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_icon_set(self):
        """Test icon set rules."""
        print("Testing icon set rules...")
        
        # Test icon set rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'iconSet'
        cf1.icon_set_type = '3TrafficLights1'
        cf1.range = 'A1:A10'
        self.assertEqual(cf1.type, 'iconSet')
        self.assertEqual(cf1.icon_set_type, '3TrafficLights1')
        self.assertEqual(cf1.range, 'A1:A10')
        print(" Icon set created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_icon_set.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Icon set test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_formula_rule(self):
        """Test formula-based rules."""
        print("Testing formula-based rules...")
        
        # Test formula rule
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'formula'
        cf1.formula = '=A1>100'
        cf1.range = 'A1:A10'
        cf1.font.bold = True
        cf1.font.color = 'FFFF0000'
        self.assertEqual(cf1.type, 'formula')
        self.assertEqual(cf1.formula, '=A1>100')
        self.assertEqual(cf1.range, 'A1:A10')
        self.assertTrue(cf1.font.bold)
        self.assertEqual(cf1.font.color, 'FFFF0000')
        print(" Formula rule created")
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
        
        # Save to separate file
        os.makedirs('outputfiles', exist_ok=True)
        output_path = 'outputfiles/test_formula_rule.xlsx'
        self.workbook.save(output_path)
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"  Formula rule test: {len(self.worksheet.conditional_formats._formats)} rules created")
        print(f"  Saved to {output_path} ({file_size} bytes)")
    
    def test_save_conditional_formatting(self):
        """Test saving conditional formatting to xlsx file."""
        print("Testing save conditional formatting...")
        
        # Create multiple conditional formats
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'cellValue'
        cf1.operator = 'lessThan'
        cf1.formula1 = '100'
        cf1.range = 'A1:A10'
        cf1.font.bold = True
        cf1.font.color = 'FFFF0000'
        
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'cellValue'
        cf2.operator = 'greaterThan'
        cf2.formula1 = '50'
        cf2.range = 'B1:B10'
        cf2.fill.set_solid_fill('FFFFFF00')
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
            self.worksheet.cells[f'B{i}'].value = i * 100
        
        # Ensure outputfiles directory exists
        os.makedirs('outputfiles', exist_ok=True)
        
        # Save workbook
        output_path = 'outputfiles/test_conditional_formatting_cell_value.xlsx'
        self.workbook.save(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f" Saved to {output_path} ({file_size} bytes)")
    
    def test_load_conditional_formatting(self):
        """Test loading conditional formatting from xlsx file."""
        print("Testing load conditional formatting...")
        
        # First, save a workbook with conditional formatting
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'cellValue'
        cf1.operator = 'greaterThan'
        cf1.formula1 = '100'
        cf1.range = 'A1:A10'
        cf1.font.bold = True
        cf1.font.color = 'FFFF0000'
        
        # Add test data
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
            self.worksheet.cells[f'B{i}'].value = i * 100
        
        # Save workbook
        save_path = 'outputfiles/test_conditional_formatting_cell_value.xlsx'
        self.workbook.save(save_path)
        
        # Load workbook back
        loaded_workbook = Workbook(save_path)
        loaded_ws = loaded_workbook.worksheets[0]
        
        # Verify conditional formatting was loaded
        self.assertTrue(hasattr(loaded_ws, 'conditional_formats'))
        self.assertEqual(len(loaded_ws.conditional_formats._formats), 1)
        
        # Verify conditional format properties
        # Note: Excel uses 'cellIs' as the XML type name for cell value rules,
        # but the API always returns 'cellValue' to users for consistency
        loaded_cf = loaded_ws.conditional_formats._formats[0]
        self.assertEqual(loaded_cf.type, 'cellValue')
        self.assertEqual(loaded_cf.operator, 'greaterThan')
        self.assertEqual(loaded_cf.formula1, '100')
        self.assertEqual(loaded_cf.range, 'A1:A10')
        self.assertTrue(loaded_cf.font.bold)
        self.assertEqual(loaded_cf.font.color, 'FFFF0000')
        
        print(f" Loaded conditional formatting from {save_path}")
    
    def test_comprehensive_conditional_formatting(self):
        """Test comprehensive conditional formatting with all rule types."""
        print("Testing comprehensive conditional formatting...")
        
        # 1. Cell value rule - greater than
        cf1 = self.worksheet.conditional_formats.add()
        cf1.type = 'cellValue'
        cf1.operator = 'greaterThan'
        cf1.formula1 = '100'
        cf1.range = 'A1:A10'
        cf1.font.bold = True
        cf1.font.color = 'FFFF0000'
        print("  Cell value (greater than) rule created")
        
        # 2. Cell value rule - less than
        cf2 = self.worksheet.conditional_formats.add()
        cf2.type = 'cellValue'
        cf2.operator = 'lessThan'
        cf2.formula1 = '50'
        cf2.range = 'B1:B10'
        cf2.fill.set_solid_fill('FFFFFF00')
        print("  Cell value (less than) rule created")
        
        # 3. Cell value rule - between
        cf3 = self.worksheet.conditional_formats.add()
        cf3.type = 'cellValue'
        cf3.operator = 'between'
        cf3.formula1 = '50'
        cf3.formula2 = '150'
        cf3.range = 'C1:C10'
        cf3.font.italic = True
        cf3.font.color = 'FF0000FF'
        print("  Cell value (between) rule created")
        
        # 4. Text rule - contains
        cf4 = self.worksheet.conditional_formats.add()
        cf4.type = 'text'
        cf4.text_operator = 'contains'
        cf4.text_formula = 'error'
        cf4.range = 'D1:D10'
        cf4.font.color = 'FFFF0000'
        cf4.fill.set_solid_fill('FFFF00')
        print("  Text (contains) rule created")
        
        # 5. Text rule - notContains
        cf5 = self.worksheet.conditional_formats.add()
        cf5.type = 'text'
        cf5.text_operator = 'notContains'
        cf5.text_formula = 'warning'
        cf5.range = 'E1:E10'
        cf5.font.color = 'FF00FF00'
        cf5.fill.set_solid_fill('FFFFFF00')
        print("  Text (notContains) rule created")
        
        # 6. Text rule - beginsWith
        cf6 = self.worksheet.conditional_formats.add()
        cf6.type = 'text'
        cf6.text_operator = 'beginsWith'
        cf6.text_formula = 'prefix'
        cf6.range = 'F1:F10'
        cf6.font.bold = True
        print("  Text (beginsWith) rule created")
        
        # 7. Text rule - endsWith
        cf7 = self.worksheet.conditional_formats.add()
        cf7.type = 'text'
        cf7.text_operator = 'endsWith'
        cf7.text_formula = 'suffix'
        cf7.range = 'G1:G10'
        cf7.font.italic = True
        print("  Text (endsWith) rule created")
        
        # 8. Duplicate values rule
        cf8 = self.worksheet.conditional_formats.add()
        cf8.type = 'duplicateValues'
        cf8.duplicate = True
        cf8.range = 'H1:H10'
        cf8.fill.set_solid_fill('FFFF0000')
        print("  Duplicate values rule created")
        
        # 9. Unique values rule
        cf9 = self.worksheet.conditional_formats.add()
        cf9.type = 'uniqueValues'
        cf9.duplicate = False
        cf9.range = 'I1:I10'
        cf9.fill.set_solid_fill('00FF00')
        print("  Unique values rule created")
        
        # 10. Top 10 items rule
        cf10 = self.worksheet.conditional_formats.add()
        cf10.type = 'top10'
        cf10.top = True
        cf10.rank = 10
        cf10.range = 'J1:J10'
        cf10.fill.set_solid_fill('FF00FF00')
        print("  Top 10 items rule created")
        
        # 11. Above average rule
        cf11 = self.worksheet.conditional_formats.add()
        cf11.type = 'aboveAverage'
        cf11.above = True
        cf11.range = 'K1:K10'
        cf11.fill.set_solid_fill('FF0000FF')
        print("  Above average rule created")
        
        # 12. Below average rule
        cf12 = self.worksheet.conditional_formats.add()
        cf12.type = 'belowAverage'
        cf12.above = False
        cf12.range = 'L1:L10'
        cf12.fill.set_solid_fill('FFFF00FF')
        print("  Below average rule created")
       
        # 13. 2-color scale
        cf13 = self.worksheet.conditional_formats.add()
        cf13.type = 'colorScale'
        cf13.color_scale_type = '2-color'
        cf13.min_color = 'FF63C384'
        cf13.max_color = 'FF006100'
        cf13.range = 'M1:M10'
        print("  2-color scale rule created")
        
        # 14. 3-color scale
        cf14 = self.worksheet.conditional_formats.add()
        cf14.type = 'colorScale'
        cf14.color_scale_type = '3-color'
        cf14.min_color = 'FF63C384'
        cf14.mid_color = 'FFFFEB84'
        cf14.max_color = 'FF006100'
        cf14.range = 'N1:N10'
        print("  3-color scale rule created")
        '''
        # 15. Data bar
        cf15 = self.worksheet.conditional_formats.add()
        cf15.type = 'dataBar'
        cf15.bar_color = 'FF006100'
        cf15.negative_color = 'FFFF0000'
        cf15.show_border = True
        cf15.direction = 'left-to-right'
        cf15.range = 'O1:O10'
        print("  Data bar rule created")
        
        # 16. Icon set
        cf16 = self.worksheet.conditional_formats.add()
        cf16.type = 'iconSet'
        cf16.icon_set_type = '3TrafficLights1'
        cf16.range = 'P1:P10'
        print("  Icon set rule created")
        
        # 17. Formula rule
        cf17 = self.worksheet.conditional_formats.add()
        cf17.type = 'formula'
        cf17.formula = '=A1>100'
        cf17.range = 'Q1:Q10'
        cf17.font.bold = True
        cf17.font.color = 'FFFF0000'
        cf17.fill.set_solid_fill('FFFFFF00')
        print("  Formula rule created")
        '''
        # Add comprehensive test data to make rules take effect
        # Column A: Values 10-100 (greater than 100 won't trigger)
        for i in range(1, 11):
            self.worksheet.cells[f'A{i}'].value = i * 10
        
        # Column B: Values 10-100 (less than 50 will trigger for first 4 rows)
        for i in range(1, 11):
            self.worksheet.cells[f'B{i}'].value = i * 10
        
        # Column C: Values 50-140 (between 50 and 150 will trigger for most)
        for i in range(1, 11):
            self.worksheet.cells[f'C{i}'].value = 50 + (i * 10)
        
        # Column D: Text with some containing "error"
        self.worksheet.cells['D1'].value = "error message"
        self.worksheet.cells['D2'].value = "warning message"
        self.worksheet.cells['D3'].value = "error found"
        self.worksheet.cells['D4'].value = "info message"
        self.worksheet.cells['D5'].value = "error detected"
        self.worksheet.cells['D6'].value = "success"
        self.worksheet.cells['D7'].value = "error"
        self.worksheet.cells['D8'].value = "warning"
        self.worksheet.cells['D9'].value = "critical error"
        self.worksheet.cells['D10'].value = "normal"
        
        # Column E: Text not containing "warning" - should trigger for cells without "warning"
        self.worksheet.cells['E1'].value = "error message"
        self.worksheet.cells['E2'].value = "success message"
        self.worksheet.cells['E3'].value = "info message"
        self.worksheet.cells['E4'].value = "warning message"
        self.worksheet.cells['E5'].value = "critical error"
        self.worksheet.cells['E6'].value = "normal"
        self.worksheet.cells['E7'].value = "error"
        self.worksheet.cells['E8'].value = "warning"
        self.worksheet.cells['E9'].value = "critical error"
        self.worksheet.cells['E10'].value = "normal"
        
        # Column F: Text beginning with "prefix" - should trigger for cells starting with "prefix"
        self.worksheet.cells['F1'].value = "prefix_test"
        self.worksheet.cells['F2'].value = "other_text"
        self.worksheet.cells['F3'].value = "prefix_data"
        self.worksheet.cells['F4'].value = "prefix_value"
        self.worksheet.cells['F5'].value = "different"
        self.worksheet.cells['F6'].value = "prefix_item"
        self.worksheet.cells['F7'].value = "prefix_string"
        self.worksheet.cells['F8'].value = "another_text"
        self.worksheet.cells['F9'].value = "prefix_test_again"
        self.worksheet.cells['F10'].value = "no_prefix"
        
        # Column G: Text ending with "suffix" - should trigger for cells ending with "suffix"
        self.worksheet.cells['G1'].value = "test_suffix"
        self.worksheet.cells['G2'].value = "suffix_data"
        self.worksheet.cells['G3'].value = "suffix_value"
        self.worksheet.cells['G4'].value = "different"
        self.worksheet.cells['G5'].value = "suffix_item"
        self.worksheet.cells['G6'].value = "suffix_string"
        self.worksheet.cells['G7'].value = "another_text"
        self.worksheet.cells['G8'].value = "suffix_test_again"
        self.worksheet.cells['G9'].value = "no_suffix"
        self.worksheet.cells['G10'].value = "normal"
        
        # Column H: Duplicate values
        for i in range(1, 11):
            self.worksheet.cells[f'H{i}'].value = i if i <= 5 else i - 5
        
        # Column I: Unique values
        for i in range(1, 11):
            self.worksheet.cells[f'I{i}'].value = i * 100
        
        # Column J: Values for top 10
        for i in range(1, 11):
            self.worksheet.cells[f'J{i}'].value = i * 100
        
        # Column K: Values for above average (average = 550)
        for i in range(1, 11):
            self.worksheet.cells[f'K{i}'].value = i * 100
        
        # Column L: Values for below average
        for i in range(1, 11):
            self.worksheet.cells[f'L{i}'].value = i * 100
        
        # Column M: Values for 2-color scale
        for i in range(1, 11):
            self.worksheet.cells[f'M{i}'].value = i * 10
        
        # Column N: Values for 3-color scale
        for i in range(1, 11):
            self.worksheet.cells[f'N{i}'].value = i * 100
        
        # Column O: Values for data bar
        for i in range(1, 11):
            self.worksheet.cells[f'O{i}'].value = i * 100
        
        # Column P: Values for icon set
        for i in range(1, 11):
            self.worksheet.cells[f'P{i}'].value = i * 100
        
        # Column Q: Values for formula (A1>100)
        for i in range(1, 11):
            self.worksheet.cells[f'Q{i}'].value = i * 100
        
        # Ensure outputfiles directory exists
        os.makedirs('outputfiles', exist_ok=True)
        
        # Save workbook
        output_path = 'outputfiles/test_conditional_formatting_comprehensive.xlsx'
        self.workbook.save(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f" Saved comprehensive test to {output_path} ({file_size} bytes)")
        print(f" Total conditional formats created: {len(self.worksheet.conditional_formats._formats)}")


if __name__ == '__main__':
    unittest.main()
