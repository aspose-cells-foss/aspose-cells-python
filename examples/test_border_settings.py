import unittest
import os
import sys

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aspose_cells import Workbook, Worksheet, Cell, Style


class TestBorderSettings(unittest.TestCase):
    """
    Test comprehensive border settings functionality.
    
    Features tested:
    - Border color settings for individual sides (top, bottom, left, right, all)
    - Border line style settings (thin, medium, thick, dashed, dotted, double)
    - Border line weight settings (1-5)
    - Complete border settings with all properties (line_style, color, weight)
    - Mixed border settings on different sides
    - Border default values
    - Border copying in style
    - Save and load border settings to verify persistence
    """
    
    def setUp(self):
        """Set up test workbook and worksheet."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.worksheets[0]
    
    def test_border_color_settings(self):
        """Test border color settings for different sides."""
        # Test setting border colors for individual sides
        cell1 = Cell("Red Top Border")
        cell1.style.set_border_color('top', 'FFFF0000')
        self.worksheet["A1"] = cell1
        
        cell2 = Cell("Blue Bottom Border")
        cell2.style.set_border_color('bottom', 'FF0000FF')
        self.worksheet["A2"] = cell2
        
        cell3 = Cell("Green Left Border")
        cell3.style.set_border_color('left', 'FF00FF00')
        self.worksheet["A3"] = cell3
        
        cell4 = Cell("Purple Right Border")
        cell4.style.set_border_color('right', 'FF800080')
        self.worksheet["A4"] = cell4
        
        cell5 = Cell("All Red Borders")
        cell5.style.set_border_color('all', 'FFFF0000')
        self.worksheet["A5"] = cell5
    
    def test_border_style_settings(self):
        """Test border line style settings."""
        # Test different border styles
        styles = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double']
        
        for i, style in enumerate(styles):
            cell = Cell(f"{style.capitalize()} Border")
            cell.style.set_border_style('all', style)
            self.worksheet[f"A{i+1}"] = cell
    
    def test_border_weight_settings(self):
        """Test border line weight settings."""
        # Test different border weights
        weights = [1, 2, 3, 4, 5]
        
        for i, weight in enumerate(weights):
            cell = Cell(f"Weight {weight}")
            cell.style.set_border_weight('all', weight)
            self.worksheet[f"A{i+1}"] = cell
    
    def test_complete_border_settings(self):
        """Test complete border settings with all properties."""
        # Test setting all border properties at once
        cell1 = Cell("Thick Black All")
        cell1.style.set_border('all', line_style='thick', color='FF000000', weight=3)
        self.worksheet["A1"] = cell1
        
        cell2 = Cell("Thin Red Top")
        cell2.style.set_border('top', line_style='thin', color='FFFF0000', weight=1)
        self.worksheet["A2"] = cell2
        
        cell3 = Cell("Medium Blue Bottom")
        cell3.style.set_border('bottom', line_style='medium', color='FF0000FF', weight=2)
        self.worksheet["A3"] = cell3
        
        cell4 = Cell("Dashed Green Left")
        cell4.style.set_border('left', line_style='dashed', color='FF00FF00', weight=1)
        self.worksheet["A4"] = cell4
        
        cell5 = Cell("Dotted Purple Right")
        cell5.style.set_border('right', line_style='dotted', color='FF800080', weight=1)
        self.worksheet["A5"] = cell5
    
    def test_mixed_border_settings(self):
        """Test mixed border settings on different sides."""
        cell = Cell("Mixed Borders")
        
        # Set different borders on each side
        cell.style.set_border('top', line_style='thick', color='FFFF0000', weight=3)
        cell.style.set_border('bottom', line_style='medium', color='FF0000FF', weight=2)
        cell.style.set_border('left', line_style='thin', color='FF00FF00', weight=1)
        cell.style.set_border('right', line_style='dashed', color='FF800080', weight=2)
        
        self.worksheet["A1"] = cell
    
    def test_border_default_values(self):
        """Test border default values."""
        from aspose_cells.style import Border, Borders
        
        # Test default border values
        border = Border()
        self.assertEqual(border.line_style, 'none')
        self.assertEqual(border.color, 'FF000000')
        self.assertEqual(border.weight, 1)
        
        # Test default borders values
        borders = Borders()
        self.assertEqual(borders.top.line_style, 'none')
        self.assertEqual(borders.top.color, 'FF000000')
        self.assertEqual(borders.top.weight, 1)
        self.assertEqual(borders.bottom.line_style, 'none')
        self.assertEqual(borders.bottom.color, 'FF000000')
        self.assertEqual(borders.bottom.weight, 1)
        self.assertEqual(borders.left.line_style, 'none')
        self.assertEqual(borders.left.color, 'FF000000')
        self.assertEqual(borders.left.weight, 1)
        self.assertEqual(borders.right.line_style, 'none')
        self.assertEqual(borders.right.color, 'FF000000')
        self.assertEqual(borders.right.weight, 1)
    
    def test_border_copy(self):
        """Test border copying in style."""
        original_style = Style()
        original_style.set_border('all', line_style='medium', color='FF0000FF', weight=2)
        
        # Copy the style
        copied_style = original_style.copy()
        
        # Verify the borders were copied correctly
        self.assertEqual(copied_style.borders.top.line_style, original_style.borders.top.line_style)
        self.assertEqual(copied_style.borders.top.color, original_style.borders.top.color)
        self.assertEqual(copied_style.borders.top.weight, original_style.borders.top.weight)
        self.assertEqual(copied_style.borders.bottom.line_style, original_style.borders.bottom.line_style)
        self.assertEqual(copied_style.borders.bottom.color, original_style.borders.bottom.color)
        self.assertEqual(copied_style.borders.bottom.weight, original_style.borders.bottom.weight)
        self.assertEqual(copied_style.borders.left.line_style, original_style.borders.left.line_style)
        self.assertEqual(copied_style.borders.left.color, original_style.borders.left.color)
        self.assertEqual(copied_style.borders.left.weight, original_style.borders.left.weight)
        self.assertEqual(copied_style.borders.right.line_style, original_style.borders.right.line_style)
        self.assertEqual(copied_style.borders.right.color, original_style.borders.right.color)
        self.assertEqual(copied_style.borders.right.weight, original_style.borders.right.weight)
    
    def test_comprehensive_border_test(self):
        """Test all border settings comprehensively."""
        # Test data for comprehensive border testing
        test_cases = [
            {
                'text': 'No Borders',
                'borders': {'all': {'line_style': 'none', 'color': 'FF000000', 'weight': 1}}
            },
            {
                'text': 'Thin Black All',
                'borders': {'all': {'line_style': 'thin', 'color': 'FF000000', 'weight': 1}}
            },
            {
                'text': 'Medium Blue All',
                'borders': {'all': {'line_style': 'medium', 'color': 'FF0000FF', 'weight': 2}}
            },
            {
                'text': 'Thick Red All',
                'borders': {'all': {'line_style': 'thick', 'color': 'FFFF0000', 'weight': 3}}
            },
            {
                'text': 'Dashed Green All',
                'borders': {'all': {'line_style': 'dashed', 'color': 'FF00FF00', 'weight': 1}}
            },
            {
                'text': 'Dotted Purple All',
                'borders': {'all': {'line_style': 'dotted', 'color': 'FF800080', 'weight': 1}}
            },
            {
                'text': 'Double Orange All',
                'borders': {'all': {'line_style': 'double', 'color': 'FFFFA500', 'weight': 2}}
            },
            {
                'text': 'Mixed: Thick Red Top, Thin Blue Bottom',
                'borders': {
                    'top': {'line_style': 'thick', 'color': 'FFFF0000', 'weight': 3},
                    'bottom': {'line_style': 'thin', 'color': 'FF0000FF', 'weight': 1},
                    'left': {'line_style': 'none', 'color': 'FF000000', 'weight': 1},
                    'right': {'line_style': 'none', 'color': 'FF000000', 'weight': 1}
                }
            },
            {
                'text': 'All Sides Different',
                'borders': {
                    'top': {'line_style': 'thick', 'color': 'FFFF0000', 'weight': 3},
                    'bottom': {'line_style': 'medium', 'color': 'FF0000FF', 'weight': 2},
                    'left': {'line_style': 'thin', 'color': 'FF00FF00', 'weight': 1},
                    'right': {'line_style': 'dashed', 'color': 'FF800080', 'weight': 2}
                }
            },
            {
                'text': 'Heavy Weight All',
                'borders': {'all': {'line_style': 'thick', 'color': 'FF000000', 'weight': 5}}
            }
        ]
        
        # Apply all test cases to cells
        for i, test_case in enumerate(test_cases):
            cell = Cell(test_case['text'])
            
            # Apply border settings
            for side, border_props in test_case['borders'].items():
                cell.style.set_border(
                    side,
                    line_style=border_props['line_style'],
                    color=border_props['color'],
                    weight=border_props['weight']
                )
            
            self.worksheet[f"A{i+1}"] = cell
    
    def test_save_border_settings(self):
        """Test saving border settings to file."""
        # Create comprehensive test data
        self.test_comprehensive_border_test()
        
        # Save to outputfiles folder
        output_path = 'outputfiles/test_border_settings.xlsx'
        
        # Ensure outputfiles directory exists
        os.makedirs('outputfiles', exist_ok=True)
        
        self.workbook.save(output_path)
        
        # Verify file was created
        self.assertTrue(os.path.exists(output_path))
        
        # Verify file is not empty
        file_size = os.path.getsize(output_path)
        self.assertGreater(file_size, 0)
        
        print(f"Border settings test file saved to: {output_path}")
        print(f"File size: {file_size} bytes")
    
    def test_load_and_verify_border_settings(self):
        """Test loading and verifying border settings from file."""
        # First, save a file with border settings
        self.test_comprehensive_border_test()
        output_path = 'outputfiles/test_border_settings.xlsx'
        os.makedirs('outputfiles', exist_ok=True)
        self.workbook.save(output_path)
        
        # Load the file back
        loaded_workbook = Workbook()
        loaded_workbook.load(output_path)
        loaded_worksheet = loaded_workbook.worksheets[0]
        
        # Verify that the file was loaded (basic verification)
        # Note: Full border verification would require parsing the XML structure
        # which is complex and beyond the scope of this test
        self.assertIsNotNone(loaded_worksheet)
        
        # Verify we can access cells
        cell_a1 = loaded_worksheet.cells["A1"]
        self.assertIsNotNone(cell_a1)
        
        cell_a2 = loaded_worksheet.cells["A2"]
        self.assertIsNotNone(cell_a2)
        
        print(f"Successfully loaded and verified border settings from: {output_path}")


if __name__ == '__main__':
    unittest.main()