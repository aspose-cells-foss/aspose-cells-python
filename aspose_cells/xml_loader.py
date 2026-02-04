"""
Aspose.Cells for Python - XML Loader Module

This module provides XML loading functionality for Excel workbooks.
It handles parsing of workbook XML files and loading data into Workbook objects.

Compatible with Aspose.Cells for .NET API structure.
ECMA-376 Compliant cell value import.
"""

import xml.etree.ElementTree as ET
from .cell_value_handler import CellValueHandler
from .comment_xml import CommentXMLReader
from .xml_autofilter_loader import AutoFilterXMLLoader
from .xml_conditional_format_loader import ConditionalFormatXMLLoader
from .xml_properties_loader import WorkbookPropertiesXMLLoader, WorksheetPropertiesXMLLoader
from .xml_hyperlink_handler import HyperlinkXMLLoader
from .xml_datavalidation_loader import DataValidationXmlLoader


class XMLLoader:
    """
    Handles loading of Excel workbook XML files.
    
    This class provides methods to parse various XML components of an Excel workbook
    including workbook structure, shared strings, styles, and worksheet data.
    """
    
    def __init__(self, workbook):
        """
        Initializes the XML loader with a workbook instance.
        
        Args:
            workbook (Workbook): The workbook instance to load data into.
        """
        self.workbook = workbook
        self.ns = {
            'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        # Initialize comment reader
        self._comment_reader = CommentXMLReader()

        # Initialize autofilter loader
        self._autofilter_loader = AutoFilterXMLLoader(self.ns)

        # Initialize conditional format loader
        self._cf_loader = ConditionalFormatXMLLoader(self.ns, workbook)

        # Initialize hyperlink loader
        self._hyperlink_loader = HyperlinkXMLLoader(self.ns)

        # Initialize data validation loader
        self._dv_loader = DataValidationXmlLoader(self.ns['main'])

        # Initialize properties loaders
        self._wb_props_loader = WorkbookPropertiesXMLLoader(self.ns)
        self._ws_props_loader = WorksheetPropertiesXMLLoader(self.ns)

    def load_workbook(self, zipf):
        """
        Loads workbook data from a ZIP file.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        # Load workbook.xml to get worksheet information
        workbook_xml_content = zipf.read('xl/workbook.xml')
        workbook_root = ET.fromstring(workbook_xml_content)

        # Load workbook properties
        self._load_workbook_properties(workbook_root)

        # Load document properties (docProps/core.xml and docProps/app.xml)
        self._load_document_properties(zipf)

        # Load worksheet information
        self._load_worksheet_info(workbook_root)

        # Load shared strings
        self._load_shared_strings(zipf)

        # Load styles
        self._load_styles(zipf)

        # Load worksheet data
        self._load_worksheets_data(zipf)
    
    def _load_workbook_properties(self, workbook_root):
        """
        Loads workbook properties from workbook XML.

        Args:
            workbook_root: The XML root element of workbook.xml.
        """
        props = self.workbook.properties

        # Load file version
        self._wb_props_loader.load_file_version(props.file_version, workbook_root)

        # Load workbook properties
        self._wb_props_loader.load_workbook_pr(props.workbook_pr, workbook_root)

        # Load workbook protection
        self._wb_props_loader.load_workbook_protection(props.protection, workbook_root)

        # Load book views
        self._wb_props_loader.load_book_views(props.view, workbook_root)

        # Load calculation properties
        self._wb_props_loader.load_calc_pr(props.calculation, workbook_root)

        # Load defined names
        self._wb_props_loader.load_defined_names(props.defined_names, workbook_root)

    def _load_worksheet_properties(self, worksheet, worksheet_root):
        """
        Loads worksheet properties from worksheet XML.

        Args:
            worksheet (Worksheet): The worksheet to load properties into.
            worksheet_root: The XML root element of the worksheet.
        """
        props = worksheet.properties

        # Load sheet views (includes selection and pane)
        self._ws_props_loader.load_sheet_views(props, worksheet_root)

        # Load sheet format properties
        self._ws_props_loader.load_sheet_format_pr(props, worksheet_root)

        # Load sheet protection
        self._ws_props_loader.load_sheet_protection(props, worksheet_root)

        # Load print options
        self._ws_props_loader.load_print_options(props, worksheet_root)

        # Load page margins
        self._ws_props_loader.load_page_margins(props, worksheet_root)

        # Load page setup
        self._ws_props_loader.load_page_setup(props, worksheet_root)

        # Load header/footer
        self._ws_props_loader.load_header_footer(props, worksheet_root)

    def _load_data_validations(self, worksheet, worksheet_root):
        """
        Loads data validations from worksheet XML.

        Args:
            worksheet (Worksheet): The worksheet to load validations into.
            worksheet_root: The XML root element of the worksheet.
        """
        validations = self._dv_loader.load_data_validations(worksheet_root)
        worksheet._data_validations = validations

    def _load_worksheet_info(self, workbook_root):
        """
        Loads worksheet information from workbook XML.

        Args:
            workbook_root: The XML root element of workbook.xml.
        """
        from .worksheet import Worksheet

        sheets = workbook_root.findall('.//main:sheet', namespaces=self.ns)
        for sheet in sheets:
            sheet_name = sheet.get('name')
            worksheet = Worksheet(sheet_name)

            # Load visibility state
            state = sheet.get('state')
            if state == 'hidden':
                worksheet._visible = False
            elif state == 'veryHidden':
                worksheet._visible = 'veryHidden'

            self.workbook._worksheets.append(worksheet)
    
    def _load_shared_strings(self, zipf):
        """
        Loads shared strings from the workbook.
        
        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            shared_strings_content = zipf.read('xl/sharedStrings.xml')
            shared_strings_root = ET.fromstring(shared_strings_content)
            self.workbook._shared_strings = []
            for si in shared_strings_root.findall('.//main:si', namespaces=self.ns):
                text_parts = [
                    t.text if t.text is not None else ''
                    for t in si.findall('.//main:t', namespaces=self.ns)
                ]
                self.workbook._shared_strings.append(''.join(text_parts))
        except KeyError:
            self.workbook._shared_strings = []
    
    def _load_styles(self, zipf):
        """
        Loads styles from the workbook.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            styles_content = zipf.read('xl/styles.xml')
            styles_root = ET.fromstring(styles_content)
            self._load_styles_xml(styles_root)
            # Load differential formatting (dxf) for conditional formatting
            self._load_dxf_styles(styles_root)
        except KeyError:
            # Use default styles
            from .xml_saver import XMLSaver
            saver = XMLSaver(self.workbook)
            saver.register_default_styles()
            self.workbook._dxf_styles = []
    
    def _load_worksheets_data(self, zipf):
        """
        Loads data for all worksheets.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        for i, worksheet in enumerate(self.workbook._worksheets):
            try:
                worksheet_content = zipf.read(f'xl/worksheets/sheet{i+1}.xml')
                worksheet_root = ET.fromstring(worksheet_content)
                self._load_worksheet_data(worksheet, worksheet_root)

                # Load comments for this worksheet
                self._comment_reader.load_comments(zipf, worksheet, i+1)

                # Load hyperlinks for this worksheet
                self._hyperlink_loader.load_hyperlinks(worksheet, worksheet_root, zipf, i+1)
            except KeyError:
                # Worksheet file not found, skip
                pass
    
    def _load_worksheet_data(self, worksheet, worksheet_root):
        """
        Loads cell data from worksheet XML according to ECMA-376 specification.
        
        Args:
            worksheet (Worksheet): The worksheet object to load data into.
            worksheet_root: The XML root element of the worksheet.
        """
        # Load used range (dimension)
        dim_elem = worksheet_root.find('main:dimension', namespaces=self.ns)
        if dim_elem is not None:
            ref = dim_elem.get('ref')
            if ref:
                from .cells import Cells
                if ':' in ref:
                    start_ref, end_ref = ref.split(':', 1)
                else:
                    start_ref = end_ref = ref
                try:
                    min_row, min_col = Cells.coordinate_from_string(start_ref)
                    max_row, max_col = Cells.coordinate_from_string(end_ref)
                    worksheet._dimension = (min_row, min_col, max_row, max_col)
                except ValueError:
                    pass

        # Load worksheet properties
        self._load_worksheet_properties(worksheet, worksheet_root)

        # Load column widths and row heights
        self._load_column_dimensions(worksheet, worksheet_root)
        self._load_row_heights(worksheet, worksheet_root)

        # Load auto filter settings (ECMA-376 Section 18.3.1.2)
        self._autofilter_loader.load_auto_filter(worksheet, worksheet_root)

        # Load conditional formatting (ECMA-376 Section 18.3.1.18)
        self._cf_loader.load_conditional_formatting(worksheet, worksheet_root)

        # Load data validations (ECMA-376 Section 18.3.1.30, 18.3.1.31)
        self._load_data_validations(worksheet, worksheet_root)

        # Find shared string table reference
        shared_strings = self.workbook._shared_strings
        
        # Load cell data
        for row_elem in worksheet_root.findall('.//main:row', namespaces=self.ns):
            for cell_elem in row_elem.findall('main:c', namespaces=self.ns):
                cell_ref = cell_elem.get('r')
                cell_type = cell_elem.get('t', 'n')  # Default to numeric per ECMA-376
                
                # Check for formula first (ECMA-376: formula must come before value)
                f_elem = cell_elem.find('main:f', namespaces=self.ns)
                formula = f_elem.text if f_elem is not None else None
                # Add '=' prefix to formula if not present (ECMA-376 stores formulas without '=')
                if formula is not None and not formula.startswith('='):
                    formula = '=' + formula
                
                # Get cell style index
                s_elem = cell_elem.get('s')
                style_idx = int(s_elem) if s_elem is not None else 0
                
                # Get cell value using CellValueHandler for ECMA-376 compliance
                v_elem = cell_elem.find('main:v', namespaces=self.ns)
                value = None
                
                if v_elem is not None and v_elem.text is not None:
                    value_str = v_elem.text
                    # Use CellValueHandler to parse value according to ECMA-376
                    value = CellValueHandler.parse_value_from_xml(
                        value_str,
                        cell_type,
                        shared_strings
                    )
                
                # Create cell with value and formula
                from .cell import Cell
                cell = Cell(value, formula)
                
                # Apply style if present
                if style_idx > 0:
                    self._apply_cell_style(cell, style_idx)
                
                # Set cell value
                worksheet.cells[cell_ref] = cell

    def _load_column_dimensions(self, worksheet, worksheet_root):
        """
        Loads column width settings from worksheet XML.
        """
        cols_elem = worksheet_root.find('main:cols', namespaces=self.ns)
        if cols_elem is None:
            return

        if not hasattr(worksheet, '_column_widths'):
            worksheet._column_widths = {}
        if not hasattr(worksheet, '_hidden_columns'):
            worksheet._hidden_columns = set()

        for col_elem in cols_elem.findall('main:col', namespaces=self.ns):
            min_val = col_elem.get('min')
            max_val = col_elem.get('max')
            width_val = col_elem.get('width')
            hidden_val = col_elem.get('hidden')
            if min_val is None or max_val is None:
                raise ValueError("Invalid column definition: missing min or max")
            try:
                min_col = int(min_val)
                max_col = int(max_val)
            except ValueError as exc:
                raise ValueError("Invalid column definition values") from exc
            if min_col < 1 or max_col < min_col:
                raise ValueError("Invalid column definition range")
            width = None
            if width_val is not None:
                try:
                    width = float(width_val)
                except ValueError as exc:
                    raise ValueError("Invalid column width value") from exc
                if width <= 0:
                    raise ValueError("Column width must be > 0")

            for col_idx in range(min_col, max_col + 1):
                if width is not None:
                    worksheet._column_widths[col_idx] = width
                if hidden_val in ('1', 'true', 'True'):
                    worksheet._hidden_columns.add(col_idx)

    def _load_row_heights(self, worksheet, worksheet_root):
        """
        Loads row height settings from worksheet XML.
        """
        if not hasattr(worksheet, '_row_heights'):
            worksheet._row_heights = {}
        if not hasattr(worksheet, '_hidden_rows'):
            worksheet._hidden_rows = set()

        for row_elem in worksheet_root.findall('.//main:row', namespaces=self.ns):
            ht = row_elem.get('ht')
            hidden_val = row_elem.get('hidden')
            if ht is None:
                if hidden_val not in ('1', 'true', 'True'):
                    continue
            row_num = row_elem.get('r')
            if row_num is None:
                raise ValueError("Row definition missing row index")
            try:
                row_idx = int(row_num)
            except ValueError as exc:
                raise ValueError("Invalid row height definition values") from exc
            if row_idx < 1:
                raise ValueError("Row index must be >= 1")
            if ht is not None:
                try:
                    height = float(ht)
                except ValueError as exc:
                    raise ValueError("Invalid row height value") from exc
                if height <= 0:
                    raise ValueError("Row height must be > 0")
                worksheet._row_heights[row_idx] = height
            if hidden_val in ('1', 'true', 'True'):
                worksheet._hidden_rows.add(row_idx)

    def _apply_cell_style(self, cell, style_idx):
        """
        Applies a style to a cell based on style index.
        
        Args:
            cell (Cell): The cell to apply style to.
            style_idx (int): The style index to apply.
        """
        cell_style_key = None
        if hasattr(self.workbook, '_cell_xf_by_index'):
            cell_style_key = self.workbook._cell_xf_by_index.get(style_idx)
        if cell_style_key is None:
            for style_key, cell_style_idx in self.workbook._cell_styles.items():
                if cell_style_idx == style_idx:
                    cell_style_key = style_key
                    break

        if cell_style_key is None:
            return

        font_key, fill_key, border_key, num_fmt_key, alignment_key, protection_key = cell_style_key

        # Apply font
        if font_key in self.workbook._font_styles:
            font_data = self.workbook._font_styles[font_key]
            cell.style.font.name = font_data['name']
            cell.style.font.size = font_data['size']
            cell.style.font.color = font_data['color']
            cell.style.font.bold = font_data['bold']
            cell.style.font.italic = font_data['italic']
            cell.style.font.underline = font_data['underline']
            cell.style.font.strikethrough = font_data['strikethrough']

        # Apply fill
        if fill_key in self.workbook._fill_styles:
            fill_data = self.workbook._fill_styles[fill_key]
            cell.style.fill.pattern_type = fill_data['pattern_type']
            cell.style.fill.foreground_color = fill_data['fg_color']
            cell.style.fill.background_color = fill_data['bg_color']

        # Apply border
        if border_key in self.workbook._border_styles:
            border_data = self.workbook._border_styles[border_key]
            cell.style.borders.top.line_style = border_data['top']['style']
            cell.style.borders.top.color = border_data['top']['color']
            cell.style.borders.bottom.line_style = border_data['bottom']['style']
            cell.style.borders.bottom.color = border_data['bottom']['color']
            cell.style.borders.left.line_style = border_data['left']['style']
            cell.style.borders.left.color = border_data['left']['color']
            cell.style.borders.right.line_style = border_data['right']['style']
            cell.style.borders.right.color = border_data['right']['color']

        # Apply number format
        if num_fmt_key in self.workbook._num_formats:
            cell.style.number_format = self.workbook._num_formats[num_fmt_key]

        # Apply alignment
        if alignment_key in self.workbook._alignment_styles:
            align_data = self.workbook._alignment_styles[alignment_key]
            cell.style.alignment.horizontal = align_data['horizontal']
            cell.style.alignment.vertical = align_data['vertical']
            cell.style.alignment.wrap_text = align_data['wrap_text']
            cell.style.alignment.indent = align_data['indent']
            cell.style.alignment.text_rotation = align_data['text_rotation']
            cell.style.alignment.shrink_to_fit = align_data['shrink_to_fit']
            cell.style.alignment.reading_order = align_data['reading_order']
            cell.style.alignment.relative_indent = align_data['relative_indent']

        # Apply protection
        if protection_key in self.workbook._protection_styles:
            prot_data = self.workbook._protection_styles[protection_key]
            cell.style.protection.locked = prot_data['locked']
            cell.style.protection.hidden = prot_data['hidden']

    def _load_styles_xml(self, styles_root):
        """
        Loads styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        # Register built-in number formats
        builtin_formats = {
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
            45: 'mm:ss',
            46: '[h]:mm:ss',
            47: 'mm:ss.0',
            48: '##0.0E+0',
            49: '@'
        }
        self.workbook._num_formats.update(builtin_formats)
        
        # Load custom number formats
        num_fmts = styles_root.findall('.//main:numFmt', namespaces=self.ns)
        for num_fmt in num_fmts:
            num_fmt_id = int(num_fmt.get('numFmtId'))
            format_code = num_fmt.get('formatCode')
            self.workbook._num_formats[num_fmt_id] = format_code
        
        # Load fonts
        self._load_fonts(styles_root)
        
        # Load fills
        self._load_fills(styles_root)
        
        # Load borders
        self._load_borders(styles_root)
        
        # Load cellXfs
        self._load_cell_xfs(styles_root)
    
    def _load_fonts(self, styles_root):
        """
        Loads font styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        fonts = styles_root.findall('.//main:font', namespaces=self.ns)
        for i, font_elem in enumerate(fonts):
            if i == 0:
                continue  # Skip default font
            sz_elem = font_elem.find('main:sz', namespaces=self.ns)
            color_elem = font_elem.find('main:color', namespaces=self.ns)
            name_elem = font_elem.find('main:name', namespaces=self.ns)
            b_elem = font_elem.find('main:b', namespaces=self.ns)
            i_elem = font_elem.find('main:i', namespaces=self.ns)
            u_elem = font_elem.find('main:u', namespaces=self.ns)
            strike_elem = font_elem.find('main:strike', namespaces=self.ns)
            
            font_data = {
                'name': name_elem.get('val') if name_elem is not None else 'Calibri',
                'size': int(sz_elem.get('val', 11)) if sz_elem is not None else 11,
                'color': color_elem.get('rgb', color_elem.get('theme', 'FF000000')) if color_elem is not None else 'FF000000',
                'bold': b_elem is not None,
                'italic': i_elem is not None,
                'underline': u_elem is not None,
                'strikethrough': strike_elem is not None
            }
            self.workbook._font_styles[i] = font_data
    
    def _load_fills(self, styles_root):
        """
        Loads fill styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        fills = styles_root.findall('.//main:fill', namespaces=self.ns)
        for i, fill_elem in enumerate(fills):
            if i < 2:
                continue  # Skip default fills (none and gray125)
            pattern_elem = fill_elem.find('main:patternFill', namespaces=self.ns)
            fg_color_elem = pattern_elem.find('main:fgColor', namespaces=self.ns) if pattern_elem is not None else None
            bg_color_elem = pattern_elem.find('main:bgColor', namespaces=self.ns) if pattern_elem is not None else None
            
            fill_data = {
                'pattern_type': pattern_elem.get('patternType', 'none') if pattern_elem is not None else 'none',
                'fg_color': fg_color_elem.get('rgb', 'FFFFFFFF') if fg_color_elem is not None else 'FFFFFFFF',
                'bg_color': bg_color_elem.get('rgb', 'FFFFFFFF') if bg_color_elem is not None else 'FFFFFFFF'
            }
            self.workbook._fill_styles[i] = fill_data
    
    def _load_borders(self, styles_root):
        """
        Loads border styles from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        borders = styles_root.findall('.//main:border', namespaces=self.ns)
        for i, border_elem in enumerate(borders):
            if i == 0:
                continue  # Skip default border
            left_elem = border_elem.find('main:left', namespaces=self.ns)
            right_elem = border_elem.find('main:right', namespaces=self.ns)
            top_elem = border_elem.find('main:top', namespaces=self.ns)
            bottom_elem = border_elem.find('main:bottom', namespaces=self.ns)
            
            # Load left border
            left_style = 'none'
            left_color = 'FF000000'
            if left_elem is not None:
                left_style = left_elem.get('style', 'none')
                left_color_elem = left_elem.find('main:color', namespaces=self.ns)
                if left_color_elem is not None:
                    left_color = left_color_elem.get('rgb', 'FF000000')
            
            # Load right border
            right_style = 'none'
            right_color = 'FF000000'
            if right_elem is not None:
                right_style = right_elem.get('style', 'none')
                right_color_elem = right_elem.find('main:color', namespaces=self.ns)
                if right_color_elem is not None:
                    right_color = right_color_elem.get('rgb', 'FF000000')
            
            # Load top border
            top_style = 'none'
            top_color = 'FF000000'
            if top_elem is not None:
                top_style = top_elem.get('style', 'none')
                top_color_elem = top_elem.find('main:color', namespaces=self.ns)
                if top_color_elem is not None:
                    top_color = top_color_elem.get('rgb', 'FF000000')
            
            # Load bottom border
            bottom_style = 'none'
            bottom_color = 'FF000000'
            if bottom_elem is not None:
                bottom_style = bottom_elem.get('style', 'none')
                bottom_color_elem = bottom_elem.find('main:color', namespaces=self.ns)
                if bottom_color_elem is not None:
                    bottom_color = bottom_color_elem.get('rgb', 'FF000000')
            
            border_data = {
                'top': {'style': top_style, 'color': top_color},
                'bottom': {'style': bottom_style, 'color': bottom_color},
                'left': {'style': left_style, 'color': left_color},
                'right': {'style': right_style, 'color': right_color}
            }
            self.workbook._border_styles[i] = border_data
    
    def _load_cell_xfs(self, styles_root):
        """
        Loads cell XF records from styles XML.
        
        Args:
            styles_root: The XML root element of styles.
        """
        if not hasattr(self.workbook, '_cell_xf_by_index'):
            self.workbook._cell_xf_by_index = {}

        cell_xfs = styles_root.findall('.//main:cellXfs/main:xf', namespaces=self.ns)
        for i, xf_elem in enumerate(cell_xfs):
            if i == 0:
                continue  # Skip default cellXf
            font_idx = int(xf_elem.get('fontId', 0))
            fill_idx = int(xf_elem.get('fillId', 0))
            border_idx = int(xf_elem.get('borderId', 0))
            num_fmt_idx = int(xf_elem.get('numFmtId', 0))
            
            # Load alignment if present
            alignment_idx = 0
            alignment_elem = xf_elem.find('main:alignment', namespaces=self.ns)
            if alignment_elem is not None:
                # Check if this alignment already exists
                horizontal = alignment_elem.get('horizontal', 'general')
                vertical = alignment_elem.get('vertical', 'bottom')
                text_rotation = int(alignment_elem.get('textRotation', 0))
                wrap_text = alignment_elem.get('wrapText') == '1'
                shrink_to_fit = alignment_elem.get('shrinkToFit') == '1'
                indent = int(alignment_elem.get('indent', 0))
                reading_order = int(alignment_elem.get('readingOrder', 0))
                relative_indent = int(alignment_elem.get('relativeIndent', 0))
                
                # Check if this alignment already exists
                for idx, align_data in self.workbook._alignment_styles.items():
                    if (align_data['horizontal'] == horizontal and
                        align_data['vertical'] == vertical and
                        align_data['wrap_text'] == wrap_text and
                        align_data['indent'] == indent and
                        align_data['text_rotation'] == text_rotation and
                        align_data['shrink_to_fit'] == shrink_to_fit and
                        align_data['reading_order'] == reading_order and
                        align_data['relative_indent'] == relative_indent):
                        alignment_idx = idx
                        break
                
                # If not found, create new alignment style
                if alignment_idx == 0:
                    alignment_idx = len(self.workbook._alignment_styles)
                    self.workbook._alignment_styles[alignment_idx] = {
                        'horizontal': horizontal,
                        'vertical': vertical,
                        'wrap_text': wrap_text,
                        'indent': indent,
                        'text_rotation': text_rotation,
                        'shrink_to_fit': shrink_to_fit,
                        'reading_order': reading_order,
                        'relative_indent': relative_indent
                    }

            # Load protection if present
            protection_idx = 0
            protection_elem = xf_elem.find('main:protection', namespaces=self.ns)
            if protection_elem is not None:
                # Get protection attributes (default: locked=1, hidden=0)
                locked = protection_elem.get('locked', '1') == '1'
                hidden = protection_elem.get('hidden', '0') == '1'

                # Check if this protection already exists
                for idx, prot_data in self.workbook._protection_styles.items():
                    if (prot_data['locked'] == locked and
                        prot_data['hidden'] == hidden):
                        protection_idx = idx
                        break

                # If not found, create new protection style
                if protection_idx == 0 and not (locked and not hidden):  # Skip if it's the default
                    protection_idx = len(self.workbook._protection_styles)
                    self.workbook._protection_styles[protection_idx] = {
                        'locked': locked,
                        'hidden': hidden
                    }

            cell_style_key = (font_idx, fill_idx, border_idx, num_fmt_idx, alignment_idx, protection_idx)
            self.workbook._cell_styles[cell_style_key] = i  # Store the actual cellXf index
            self.workbook._cell_xf_by_index[i] = cell_style_key

    def _load_dxf_styles(self, styles_root):
        """
        Loads differential formatting (dxf) styles from styles XML.

        These are used for conditional formatting.

        Args:
            styles_root: The XML root element of styles.
        """
        self.workbook._dxf_styles = []

        dxfs_elem = styles_root.find('.//main:dxfs', namespaces=self.ns)
        if dxfs_elem is None:
            return

        for dxf_elem in dxfs_elem.findall('main:dxf', namespaces=self.ns):
            dxf_data = {}

            # Load font
            font_elem = dxf_elem.find('main:font', namespaces=self.ns)
            if font_elem is not None:
                font_data = {}
                b_elem = font_elem.find('main:b', namespaces=self.ns)
                if b_elem is not None:
                    font_data['bold'] = b_elem.get('val', '1') != '0'
                i_elem = font_elem.find('main:i', namespaces=self.ns)
                if i_elem is not None:
                    font_data['italic'] = i_elem.get('val', '1') != '0'
                u_elem = font_elem.find('main:u', namespaces=self.ns)
                if u_elem is not None:
                    font_data['underline'] = True
                strike_elem = font_elem.find('main:strike', namespaces=self.ns)
                if strike_elem is not None:
                    font_data['strikethrough'] = True
                color_elem = font_elem.find('main:color', namespaces=self.ns)
                if color_elem is not None:
                    font_data['color'] = color_elem.get('rgb', 'FF000000')
                if font_data:
                    dxf_data['font'] = font_data

            # Load fill
            fill_elem = dxf_elem.find('main:fill', namespaces=self.ns)
            if fill_elem is not None:
                pattern_elem = fill_elem.find('main:patternFill', namespaces=self.ns)
                if pattern_elem is not None:
                    fill_data = {
                        'pattern_type': pattern_elem.get('patternType', 'solid')
                    }
                    fg_elem = pattern_elem.find('main:fgColor', namespaces=self.ns)
                    if fg_elem is not None:
                        fill_data['fg_color'] = fg_elem.get('rgb', 'FFFFFFFF')
                    bg_elem = pattern_elem.find('main:bgColor', namespaces=self.ns)
                    if bg_elem is not None:
                        fill_data['bg_color'] = bg_elem.get('rgb', 'FFFFFFFF')
                    dxf_data['fill'] = fill_data

            # Load border (simplified - just check if any border is present)
            border_elem = dxf_elem.find('main:border', namespaces=self.ns)
            if border_elem is not None:
                # Check any side for style
                for side in ['left', 'right', 'top', 'bottom']:
                    side_elem = border_elem.find(f'main:{side}', namespaces=self.ns)
                    if side_elem is not None:
                        style = side_elem.get('style', 'thin')
                        color = 'FF000000'
                        color_elem = side_elem.find('main:color', namespaces=self.ns)
                        if color_elem is not None:
                            color = color_elem.get('rgb', 'FF000000')
                        dxf_data['border'] = {'style': style, 'color': color}
                        break

            self.workbook._dxf_styles.append(dxf_data)

    def _load_document_properties(self, zipf):
        """
        Loads document properties from docProps/core.xml and docProps/app.xml.

        ECMA-376 Part 2, Section 11 - Core Properties
        ECMA-376 Part 1, Section 22.2 - Extended Properties

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        # Load core properties
        self._load_core_properties(zipf)

        # Load extended/app properties
        self._load_app_properties(zipf)

    def _load_core_properties(self, zipf):
        """
        Loads core document properties from docProps/core.xml.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            core_xml_content = zipf.read('docProps/core.xml')
            core_root = ET.fromstring(core_xml_content)

            # Namespaces for core properties
            ns = {
                'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'dc': 'http://purl.org/dc/elements/1.1/',
                'dcterms': 'http://purl.org/dc/terms/'
            }

            # Ensure document_properties exists
            if not hasattr(self.workbook, 'document_properties') or self.workbook.document_properties is None:
                from .document_properties import DocumentProperties
                self.workbook._document_properties = DocumentProperties()

            core = self.workbook.document_properties.core

            # Load Dublin Core properties
            title = core_root.find('dc:title', ns)
            if title is not None and title.text:
                core._title = title.text

            subject = core_root.find('dc:subject', ns)
            if subject is not None and subject.text:
                core._subject = subject.text

            creator = core_root.find('dc:creator', ns)
            if creator is not None and creator.text:
                core._creator = creator.text

            description = core_root.find('dc:description', ns)
            if description is not None and description.text:
                core._description = description.text

            # Load OPC Core Properties
            keywords = core_root.find('cp:keywords', ns)
            if keywords is not None and keywords.text:
                core._keywords = keywords.text

            last_modified_by = core_root.find('cp:lastModifiedBy', ns)
            if last_modified_by is not None and last_modified_by.text:
                core._last_modified_by = last_modified_by.text

            revision = core_root.find('cp:revision', ns)
            if revision is not None and revision.text:
                core._revision = revision.text

            category = core_root.find('cp:category', ns)
            if category is not None and category.text:
                core._category = category.text

            content_status = core_root.find('cp:contentStatus', ns)
            if content_status is not None and content_status.text:
                core._content_status = content_status.text

            # Load dates
            created = core_root.find('dcterms:created', ns)
            if created is not None and created.text:
                core._created = self._parse_datetime(created.text)

            modified = core_root.find('dcterms:modified', ns)
            if modified is not None and modified.text:
                core._modified = self._parse_datetime(modified.text)

        except KeyError:
            # docProps/core.xml not found, skip
            pass

    def _load_app_properties(self, zipf):
        """
        Loads extended/application properties from docProps/app.xml.

        Args:
            zipf: A ZipFile object containing the workbook data.
        """
        try:
            app_xml_content = zipf.read('docProps/app.xml')
            app_root = ET.fromstring(app_xml_content)

            # Ensure document_properties exists
            if not hasattr(self.workbook, 'document_properties') or self.workbook.document_properties is None:
                from .document_properties import DocumentProperties
                self.workbook._document_properties = DocumentProperties()

            ext = self.workbook.document_properties.extended

            # Load properties (note: default namespace, so no prefix)
            application = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Application')
            if application is not None and application.text:
                ext._application = application.text

            app_version = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}AppVersion')
            if app_version is not None and app_version.text:
                ext._app_version = app_version.text

            company = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company')
            if company is not None and company.text:
                ext._company = company.text

            manager = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Manager')
            if manager is not None and manager.text:
                ext._manager = manager.text

            hyperlink_base = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}HyperlinkBase')
            if hyperlink_base is not None and hyperlink_base.text:
                ext._hyperlink_base = hyperlink_base.text

            doc_security = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}DocSecurity')
            if doc_security is not None and doc_security.text:
                ext._doc_security = int(doc_security.text)

            scale_crop = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}ScaleCrop')
            if scale_crop is not None and scale_crop.text:
                ext._scale_crop = scale_crop.text.lower() == 'true'

            links_up_to_date = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}LinksUpToDate')
            if links_up_to_date is not None and links_up_to_date.text:
                ext._links_up_to_date = links_up_to_date.text.lower() == 'true'

            shared_doc = app_root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}SharedDoc')
            if shared_doc is not None and shared_doc.text:
                ext._shared_doc = shared_doc.text.lower() == 'true'

        except KeyError:
            # docProps/app.xml not found, skip
            pass

    def _parse_datetime(self, date_str):
        """
        Parses a W3CDTF datetime string.

        Args:
            date_str: A datetime string in W3CDTF format (e.g., '2024-01-15T10:30:00Z')

        Returns:
            datetime object or the original string if parsing fails.
        """
        from datetime import datetime

        try:
            # Try ISO format with Z suffix
            if date_str.endswith('Z'):
                return datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            # Try ISO format
            return datetime.fromisoformat(date_str)
        except (ValueError, AttributeError):
            # Return the string if parsing fails
            return date_str
