"""
Aspose.Cells for Python - XML Saver Module

This module provides the XMLSaver class which handles saving workbook data to XML format.
The XMLSaver class creates all the necessary XML files for an Excel .xlsx file.

Compatible with Aspose.Cells for .NET API structure.
ECMA-376 Compliant cell value export.
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from .cell_value_handler import CellValueHandler
from .shared_strings import SharedStringTable
from .comment_xml import CommentXMLWriter
from .xml_autofilter_saver import AutoFilterXMLWriter
from .xml_conditional_format_saver import ConditionalFormatXMLWriter
from .xml_properties_saver import WorkbookPropertiesXMLWriter, WorksheetPropertiesXMLWriter
from .xml_hyperlink_handler import HyperlinkXMLSaver, HyperlinkRelationshipWriter
from .xml_datavalidation_saver import DataValidationXmlSaver


class XMLSaver:
    """
    Handles saving workbook data to XML format for .xlsx files.
    
    The XMLSaver class is responsible for creating all of the XML files that make up
    an Excel .xlsx file, including content types, relationships, workbook, styles,
    shared strings, and worksheet files.
    
    Examples:
        >>> saver = XMLSaver(workbook)
        >>> saver.save('output.xlsx')
    """
    
    def __init__(self, workbook):
        """
        Initializes a new instance of the XMLSaver class.
        
        Args:
            workbook (Workbook): The workbook to save.
        """
        self._workbook = workbook
        
        # Initialize style dictionaries if they don't exist
        if not hasattr(workbook, '_font_styles'):
            workbook._font_styles = {}
        if not hasattr(workbook, '_fill_styles'):
            workbook._fill_styles = {}
        if not hasattr(workbook, '_border_styles'):
            workbook._border_styles = {}
        if not hasattr(workbook, '_alignment_styles'):
            workbook._alignment_styles = {}
        if not hasattr(workbook, '_cell_styles'):
            workbook._cell_styles = {}
        if not hasattr(workbook, '_num_formats'):
            workbook._num_formats = {}
        
        # Initialize shared string table
        self._shared_string_table = SharedStringTable()

        # Initialize comment writer
        self._comment_writer = CommentXMLWriter()

        # Initialize autofilter writer
        self._autofilter_writer = AutoFilterXMLWriter(self._escape_xml)

        # Initialize conditional formatting writer
        self._cf_writer = ConditionalFormatXMLWriter(self._escape_xml)

        # Initialize hyperlink writer
        self._hyperlink_writer = HyperlinkXMLSaver()

        # Initialize data validation writer
        self._dv_writer = DataValidationXmlSaver()

        # Initialize properties writers
        self._wb_props_writer = WorkbookPropertiesXMLWriter(self._escape_xml)
        self._ws_props_writer = WorksheetPropertiesXMLWriter(self._escape_xml)

        # Initialize differential formatting (dxf) collection for conditional formatting
        self._dxf_styles = []

    def _register_conditional_format_dxfs(self):
        """
        Registers differential formatting (dxf) styles for all conditional formats.

        This method assigns dxfId to each conditional format that has formatting applied.
        The dxf styles are stored in _dxf_styles for later writing to styles.xml.
        """
        self._dxf_styles = []

        for worksheet in self._workbook.worksheets:
            for cf in worksheet.conditional_formats:
                # Skip rules that don't use dxf (colorScale, dataBar, iconSet)
                if cf._type in ('colorScale', 'dataBar', 'iconSet'):
                    cf._dxf_id = None
                    continue

                # Check if this conditional format has any formatting applied
                has_formatting = self._cf_has_formatting(cf)

                if has_formatting:
                    # Create dxf entry and assign ID
                    dxf_data = self._create_dxf_data(cf)
                    cf._dxf_id = len(self._dxf_styles)
                    self._dxf_styles.append(dxf_data)
                else:
                    cf._dxf_id = None

    def _cf_has_formatting(self, cf):
        """Checks if a conditional format has any formatting applied."""
        # Check font
        if cf._font:
            if (cf._font.bold or cf._font.italic or cf._font.underline or
                cf._font.strikethrough or cf._font.color != 'FF000000'):
                return True

        # Check fill
        if cf._fill:
            if cf._fill.pattern_type != 'none' and cf._fill.foreground_color != 'FFFFFFFF':
                return True

        # Check border
        if cf._border:
            if cf._border.line_style != 'none':
                return True

        return False

    def _create_dxf_data(self, cf):
        """Creates a dxf data dictionary from a conditional format."""
        dxf_data = {}

        # Add font data if modified
        if cf._font:
            font_data = {}
            if cf._font.bold:
                font_data['bold'] = True
            if cf._font.italic:
                font_data['italic'] = True
            if cf._font.underline:
                font_data['underline'] = True
            if cf._font.strikethrough:
                font_data['strikethrough'] = True
            if cf._font.color != 'FF000000':
                font_data['color'] = cf._font.color
            if font_data:
                dxf_data['font'] = font_data

        # Add fill data if modified
        if cf._fill and cf._fill.pattern_type != 'none':
            dxf_data['fill'] = {
                'pattern_type': cf._fill.pattern_type,
                'fg_color': cf._fill.foreground_color,
                'bg_color': cf._fill.background_color
            }

        # Add border data if modified
        if cf._border and cf._border.line_style != 'none':
            dxf_data['border'] = {
                'style': cf._border.line_style,
                'color': cf._border.color
            }

        return dxf_data

    def save(self, file_path):
        """
        Saves the workbook to an Excel file (.xlsx format).
        
        Args:
            file_path (str): Path where the Excel file should be saved.
            
        Examples:
            >>> saver = XMLSaver(workbook)
            >>> saver.save('output.xlsx')
        """
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Create a ZIP file (XLSX is a ZIP archive)
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Write [Content_Types].xml
            self._write_content_types(zipf)
            
            # Write _rels/.rels
            self._write_root_relationships(zipf)
            
            # Write xl/_rels/workbook.xml.rels
            self._write_workbook_relationships(zipf)
            
            # Write xl/workbook.xml
            self._write_workbook_xml(zipf)
            
            # Register differential formats (dxf) for conditional formatting
            # This must be done BEFORE writing worksheets so cfRules can reference dxfId
            self._register_conditional_format_dxfs()

            # Process worksheets first to populate shared string table and style collections
            # This must be done BEFORE writing shared strings and styles XML
            for i, worksheet in enumerate(self._workbook.worksheets):
                # Write xl/worksheets/sheet{i+1}.xml
                self._write_worksheet_xml(zipf, worksheet, i+1)

                # Write xl/worksheets/_rels/sheet{i+1}.xml.rels
                self._write_worksheet_relationships(zipf, i+1)

                # Write xl/comments{i+1}.xml if worksheet has comments
                self._comment_writer.write_comments_xml(zipf, worksheet, i+1)

                # Write xl/drawings/vmlDrawing{i+1}.vml if worksheet has comments
                self._comment_writer.write_vml_drawing_xml(zipf, worksheet, i+1)

            # Write xl/styles.xml (AFTER processing worksheets to ensure styles are registered)
            self._write_styles_xml(zipf)
            
            # Write xl/sharedStrings.xml (AFTER processing worksheets)
            self._write_shared_strings_xml(zipf)

            # Write docProps/core.xml (document properties)
            self._write_core_properties_xml(zipf)

            # Write docProps/app.xml (extended properties)
            self._write_app_properties_xml(zipf)
    
    def _write_content_types(self, zipf):
        """Writes [Content_Types].xml file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        content += '    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        content += '    <Default Extension="xml" ContentType="application/xml"/>\n'
        content += '    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
        content += '    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n'
        content += '    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n'
        
        # Add worksheet content types
        for i in range(len(self._workbook.worksheets)):
            content += f'    <Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        
        # Add comments and VML drawing content types for worksheets that have comments
        for i, worksheet in enumerate(self._workbook.worksheets):
            if self._comment_writer.worksheet_has_comments(worksheet):
                content += f'    <Override PartName="/xl/comments{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>\n'
                content += f'    <Override PartName="/xl/drawings/vmlDrawing{i+1}.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>\n'

        # Add docProps content types
        content += '    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n'
        content += '    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n'

        content += '</Types>\n'
        zipf.writestr('[Content_Types].xml', content)
    
    def _write_root_relationships(self, zipf):
        """Writes _rels/.rels file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        content += '    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
        content += '    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\n'
        content += '    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\n'
        content += '</Relationships>\n'
        zipf.writestr('_rels/.rels', content)
    
    def _write_workbook_relationships(self, zipf):
        """Writes xl/_rels/workbook.xml.rels file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        
        # Add worksheet relationships
        for i in range(len(self._workbook.worksheets)):
            content += f'    <Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>\n'
        
        # Add styles and shared strings relationships
        content += '    <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        content += '    <Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n'
        content += '</Relationships>\n'
        zipf.writestr('xl/_rels/workbook.xml.rels', content)
    
    def _write_workbook_xml(self, zipf):
        """Writes xl/workbook.xml file."""
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'

        # Write workbook properties
        props = self._workbook.properties

        # File version (ECMA-376 Section 18.2.10)
        content += self._wb_props_writer.format_file_version_xml(props.file_version)

        # Workbook properties (ECMA-376 Section 18.2.13)
        content += self._wb_props_writer.format_workbook_pr_xml(props.workbook_pr)

        # Workbook protection (ECMA-376 Section 18.2.29)
        content += self._wb_props_writer.format_workbook_protection_xml(props.protection)

        # Book views (ECMA-376 Section 18.2.1)
        content += self._wb_props_writer.format_book_views_xml(props.view)

        # Sheets
        content += '    <sheets>\n'

        # Add sheet elements
        for i, worksheet in enumerate(self._workbook.worksheets):
            state_attr = ''
            if worksheet.visible == False:
                state_attr = ' state="hidden"'
            elif worksheet.visible == 'veryHidden':
                state_attr = ' state="veryHidden"'
            content += f'        <sheet name="{self._escape_xml(worksheet.name)}" sheetId="{i+1}"{state_attr} r:id="rId{i+1}"/>\n'

        content += '    </sheets>\n'

        # Defined names (ECMA-376 Section 18.2.6)
        content += self._wb_props_writer.format_defined_names_xml(props.defined_names)

        # Calculation properties (ECMA-376 Section 18.2.2)
        content += self._wb_props_writer.format_calc_pr_xml(props.calculation)

        content += '</workbook>\n'
        zipf.writestr('xl/workbook.xml', content)
    
    def _write_worksheet_xml(self, zipf, worksheet, sheet_num):
        """
        Writes worksheet XML file with ECMA-376 compliant cell values.
        
        ECMA-376 Part 1, Section 18.3.1.73 specifies that cells must be grouped
        by row elements within the sheetData element.
        
        Args:
            zipf: The ZIP file object to write to.
            worksheet: The worksheet object to save.
            sheet_num: The worksheet number (1-based).
        """
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'

        # Write worksheet properties
        ws_props = worksheet.properties

        # Sheet views (ECMA-376 Section 18.3.1.88)
        is_first_sheet = (sheet_num == 1)
        content += self._ws_props_writer.format_sheet_views_xml(ws_props, is_selected=is_first_sheet)

        # Sheet format properties (ECMA-376 Section 18.3.1.82)
        content += self._ws_props_writer.format_sheet_format_pr_xml(ws_props.format)

        # Write column widths/hidden columns if configured
        if getattr(worksheet, '_column_widths', None) or getattr(worksheet, '_hidden_columns', None):
            content += self._format_cols_xml(worksheet)

        content += '    <sheetData>\n'
        
        # Get all cells with their references
        cells = worksheet.cells.get_all_cells()
        
        # Sort cells by reference to ensure proper order
        sorted_refs = sorted(cells.keys(), key=self._cell_reference_sort_key)
        
        # Group cells by row (ECMA-376 requirement)
        rows = {}
        for ref in sorted_refs:
            row, col = self._cell_reference_sort_key(ref)
            if row not in rows:
                rows[row] = []
            rows[row].append((ref, cells[ref]))

        # Ensure rows with custom heights are included even if they have no cells
        if getattr(worksheet, '_row_heights', None):
            for row_num in worksheet._row_heights.keys():
                if row_num not in rows:
                    rows[row_num] = []
        if getattr(worksheet, '_hidden_rows', None):
            for row_num in worksheet._hidden_rows:
                if row_num not in rows:
                    rows[row_num] = []
        
        # Write row elements with cells (ECMA-376 compliant structure)
        for row_num in sorted(rows.keys()):
            row_attrs = [f'r="{row_num}"']
            row_height = None
            if getattr(worksheet, '_row_heights', None):
                row_height = worksheet._row_heights.get(row_num)
            if row_height is not None:
                row_attrs.append(f'ht="{row_height}"')
                row_attrs.append('customHeight="1"')
            if getattr(worksheet, '_hidden_rows', None) and row_num in worksheet._hidden_rows:
                row_attrs.append('hidden="1"')
            content += f'        <row {" ".join(row_attrs)}>\n'
            for ref, cell in rows[row_num]:
                content += self._format_cell_xml(ref, cell)
            content += '        </row>\n'
        
        content += '    </sheetData>\n'

        # Sheet protection (ECMA-376 Section 18.3.1.85)
        content += self._ws_props_writer.format_sheet_protection_xml(ws_props.protection)

        # Write auto filter settings (ECMA-376 Section 18.3.1.2)
        # autoFilter must come AFTER sheetData per ECMA-376 schema sequence
        if worksheet.auto_filter.range is not None:
            content += self._autofilter_writer.format_auto_filter_xml(worksheet.auto_filter)

        # Write conditional formatting (ECMA-376 Section 18.3.1.18)
        # conditionalFormatting must come AFTER autoFilter per ECMA-376 schema sequence
        if len(worksheet.conditional_formats) > 0:
            content += self._cf_writer.format_conditional_formatting_xml(worksheet.conditional_formats)

        # Write hyperlinks (ECMA-376 Section 18.3.1.48)
        # hyperlinks must come AFTER conditionalFormatting per ECMA-376 schema sequence
        if worksheet.hyperlinks.count > 0:
            # Reset relationship counter for this worksheet
            self._hyperlink_writer.reset_relationship_counter()
            content += self._hyperlink_writer.format_hyperlinks_xml(worksheet)

        # Write data validations (ECMA-376 Section 18.3.1.30, 18.3.1.31)
        # dataValidations must come AFTER hyperlinks per ECMA-376 schema sequence
        if worksheet.data_validations.count > 0:
            content += self._format_data_validations_xml(worksheet.data_validations)

        # Print options (ECMA-376 Section 18.3.1.70)
        content += self._ws_props_writer.format_print_options_xml(ws_props.print_options)

        # Page margins (ECMA-376 Section 18.3.1.62)
        content += self._ws_props_writer.format_page_margins_xml(ws_props.page_margins)

        # Page setup (ECMA-376 Section 18.3.1.63)
        content += self._ws_props_writer.format_page_setup_xml(ws_props.page_setup)

        # Header/footer (ECMA-376 Section 18.3.1.46)
        content += self._ws_props_writer.format_header_footer_xml(ws_props.header_footer)

        # Add legacy drawing reference if worksheet has comments
        if self._comment_writer.worksheet_has_comments(worksheet):
            content += '    <legacyDrawing r:id="rId1"/>\n'

        content += '</worksheet>\n'
        zipf.writestr(f'xl/worksheets/sheet{sheet_num}.xml', content)

    def _format_cols_xml(self, worksheet):
        """
        Formats column width settings as <cols> XML.
        """
        col_widths = getattr(worksheet, '_column_widths', None) or {}
        hidden_cols = getattr(worksheet, '_hidden_columns', None) or set()
        if not col_widths and not hidden_cols:
            return ''

        lines = ['    <cols>']
        all_cols = sorted(set(col_widths.keys()) | set(hidden_cols))
        for col_idx in all_cols:
            attrs = [f'min="{col_idx}"', f'max="{col_idx}"']
            if col_idx in col_widths:
                width = col_widths[col_idx]
                attrs.append(f'width="{width}"')
                attrs.append('customWidth="1"')
            if col_idx in hidden_cols:
                attrs.append('hidden="1"')
            lines.append(f'        <col {" ".join(attrs)}/>')
        lines.append('    </cols>\n')
        return '\n'.join(lines)
    
    def _cell_reference_sort_key(self, ref):
        """
        Creates a sort key for cell references.
        
        Args:
            ref (str): Cell reference (e.g., "A1", "B3")
            
        Returns:
            tuple: (row, column) for sorting
        """
        from .cells import Cells
        row, col = Cells.coordinate_from_string(ref)
        return (row, col)

    def _format_data_validations_xml(self, validations):
        """
        Formats data validations as XML according to ECMA-376 specification.

        Args:
            validations (DataValidationCollection): The data validations collection.

        Returns:
            str: XML string for data validations.
        """
        from .data_validation import (
            DataValidationType, DataValidationOperator,
            DataValidationAlertStyle, DataValidationImeMode
        )

        # Mapping from enum values to XML attribute values
        type_map = {
            DataValidationType.NONE: 'none',
            DataValidationType.WHOLE_NUMBER: 'whole',
            DataValidationType.DECIMAL: 'decimal',
            DataValidationType.LIST: 'list',
            DataValidationType.DATE: 'date',
            DataValidationType.TIME: 'time',
            DataValidationType.TEXT_LENGTH: 'textLength',
            DataValidationType.CUSTOM: 'custom',
        }

        operator_map = {
            DataValidationOperator.BETWEEN: 'between',
            DataValidationOperator.NOT_BETWEEN: 'notBetween',
            DataValidationOperator.EQUAL: 'equal',
            DataValidationOperator.NOT_EQUAL: 'notEqual',
            DataValidationOperator.GREATER_THAN: 'greaterThan',
            DataValidationOperator.LESS_THAN: 'lessThan',
            DataValidationOperator.GREATER_THAN_OR_EQUAL: 'greaterThanOrEqual',
            DataValidationOperator.LESS_THAN_OR_EQUAL: 'lessThanOrEqual',
        }

        alert_map = {
            DataValidationAlertStyle.STOP: 'stop',
            DataValidationAlertStyle.WARNING: 'warning',
            DataValidationAlertStyle.INFORMATION: 'information',
        }

        ime_map = {
            DataValidationImeMode.NO_CONTROL: 'noControl',
            DataValidationImeMode.OFF: 'off',
            DataValidationImeMode.ON: 'on',
            DataValidationImeMode.DISABLED: 'disabled',
            DataValidationImeMode.HIRAGANA: 'hiragana',
            DataValidationImeMode.FULL_KATAKANA: 'fullKatakana',
            DataValidationImeMode.HALF_KATAKANA: 'halfKatakana',
            DataValidationImeMode.FULL_ALPHA: 'fullAlpha',
            DataValidationImeMode.HALF_ALPHA: 'halfAlpha',
            DataValidationImeMode.FULL_HANGUL: 'fullHangul',
            DataValidationImeMode.HALF_HANGUL: 'halfHangul',
        }

        xml = f'<dataValidations count="{validations.count}"'

        if validations.disable_prompts:
            xml += ' disablePrompts="1"'
        if validations.x_window is not None:
            xml += f' xWindow="{validations.x_window}"'
        if validations.y_window is not None:
            xml += f' yWindow="{validations.y_window}"'

        xml += '>'

        for dv in validations:
            xml += '<dataValidation'

            # Required attribute: sqref
            if dv.sqref:
                xml += f' sqref="{self._escape_xml(dv.sqref)}"'

            # Type attribute (only if not default 'none')
            if dv.type != DataValidationType.NONE:
                xml += f' type="{type_map.get(dv.type, "none")}"'

            # Operator attribute (for types that use operators, only if not default)
            if dv.type in (DataValidationType.WHOLE_NUMBER, DataValidationType.DECIMAL,
                           DataValidationType.DATE, DataValidationType.TIME,
                           DataValidationType.TEXT_LENGTH):
                if dv.operator != DataValidationOperator.BETWEEN:
                    xml += f' operator="{operator_map.get(dv.operator, "between")}"'

            # Error style (only if not default 'stop')
            if dv.alert_style != DataValidationAlertStyle.STOP:
                xml += f' errorStyle="{alert_map.get(dv.alert_style, "stop")}"'

            # IME mode (only if not default 'noControl')
            if dv.ime_mode != DataValidationImeMode.NO_CONTROL:
                xml += f' imeMode="{ime_map.get(dv.ime_mode, "noControl")}"'

            # Boolean attributes
            if dv.allow_blank:
                xml += ' allowBlank="1"'

            # Note: showDropDown="1" means HIDE dropdown (counterintuitive ECMA-376 naming)
            if not dv.show_dropdown:
                xml += ' showDropDown="1"'

            if dv.show_input_message:
                xml += ' showInputMessage="1"'

            if dv.show_error_message:
                xml += ' showErrorMessage="1"'

            # String attributes
            if dv.error_title:
                xml += f' errorTitle="{self._escape_xml(dv.error_title)}"'

            if dv.error_message:
                xml += f' error="{self._escape_xml(dv.error_message)}"'

            if dv.input_title:
                xml += f' promptTitle="{self._escape_xml(dv.input_title)}"'

            if dv.input_message:
                xml += f' prompt="{self._escape_xml(dv.input_message)}"'

            xml += '>'

            # Formula elements
            if dv.formula1 is not None:
                xml += f'<formula1>{self._escape_xml(dv.formula1)}</formula1>'

            if dv.formula2 is not None:
                xml += f'<formula2>{self._escape_xml(dv.formula2)}</formula2>'

            xml += '</dataValidation>'

        xml += '</dataValidations>'
        return xml

    def _format_cell_xml(self, ref, cell):
        """
        Formats a cell as XML according to ECMA-376 specification.
        
        Args:
            ref (str): Cell reference (e.g., "A1")
            cell (Cell): The cell object
            
        Returns:
            str: XML representation of the cell
        """
        # Get or create cell style index
        style_idx = self.get_or_create_cell_style(cell)
        
        # Format value using CellValueHandler for ECMA-376 compliance
        value_str, cell_type = CellValueHandler.format_value_for_xml(cell.value)
        
        # Handle shared strings
        if cell_type == CellValueHandler.TYPE_SHARED_STRING and value_str is not None:
            # Add to shared string table and get index
            shared_string_index = self._shared_string_table.add_string(value_str)
            value_str = str(shared_string_index)
        
        # Build cell XML
        # ECMA-376: cell element with r (reference), s (style), and t (type) attributes
        # ECMA-376: formula (<f>) must come before value (<v>)
        
        if style_idx > 0 and cell_type is not None:
            xml = f'        <c r="{ref}" s="{style_idx}" t="{cell_type}">\n'
        elif style_idx > 0:
            xml = f'        <c r="{ref}" s="{style_idx}">\n'
        elif cell_type is not None:
            xml = f'        <c r="{ref}" t="{cell_type}">\n'
        else:
            xml = f'        <c r="{ref}">\n'
        
        # Add formula if present (ECMA-376: formula must come before value)
        if cell.formula:
            # Remove leading '=' from formula for ECMA-376 compliance
            formula_text = cell.formula.lstrip('=')
            escaped_formula = self._escape_xml(formula_text)
            xml += f'            <f>{escaped_formula}</f>\n'
        
        # Add value if present
        if value_str is not None:
            escaped_value = self._escape_xml(value_str)
            xml += f'            <v>{escaped_value}</v>\n'
        
        xml += '        </c>\n'
        
        return xml
    
    def _escape_xml(self, text):
        """
        Escapes special characters for XML according to ECMA-376.
        
        ECMA-376 Part 1, Section 3.2.20 specifies that the following characters
        must be escaped in XML content:
        - & (ampersand) -> &amp;
        - < (less than) -> &lt;
        - > (greater than) -> &gt;
        - " (double quote) -> &quot;
        - ' (apostrophe/single quote) -> &apos;
        
        Note: The > character only needs to be escaped when it appears in the
        sequence ]]> to avoid confusion with CDATA section end markers.
        However, it's good practice to always escape it for consistency.
        
        Args:
            text (str): The text to escape
            
        Returns:
            str: The escaped text, or None if input is None
        """
        if text is None:
            return None
        
        # Ensure we're working with a string
        if not isinstance(text, str):
            text = str(text)
        
        # Escape characters in the correct order to avoid double-escaping
        # Order matters: & must be escaped first to avoid escaping the & in other entities
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')
        
        return text
    
    def _write_worksheet_relationships(self, zipf, sheet_num):
        """Writes xl/worksheets/_rels/sheet{sheet_num}.xml.rels file."""
        worksheet = self._workbook.worksheets[sheet_num - 1]

        # Collect existing relationships (comments, VML)
        existing_rels = []
        if self._comment_writer.worksheet_has_comments(worksheet):
            existing_rels.append(('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
                                 f'../drawings/vmlDrawing{sheet_num}.vml', None))
            existing_rels.append(('rId2', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
                                 f'../comments{sheet_num}.xml', None))

        # Get hyperlink relationships
        hyperlink_rels = self._hyperlink_writer.get_hyperlink_relationships(worksheet)

        # Only write relationships file if there are relationships to write
        if existing_rels or hyperlink_rels:
            content = HyperlinkRelationshipWriter.format_relationships_xml(hyperlink_rels, existing_rels)
            zipf.writestr(f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels', content)

    def _write_styles_xml(self, zipf):
        """Writes xl/styles.xml file."""
        # Register default styles
        self.register_default_styles()
        
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        
        # Write number formats
        custom_num_fmts = {k: v for k, v in self._workbook._num_formats.items() if k >= 164}
        content += f'    <numFmts count="{len(custom_num_fmts)}">\n'
        for num_fmt_id, format_code in sorted(custom_num_fmts.items()):
            escaped_code = self._escape_xml(format_code)
            content += f'        <numFmt numFmtId="{num_fmt_id}" formatCode="{escaped_code}"/>\n'
        content += '    </numFmts>\n'
        
        # Write fonts
        content += f'    <fonts count="{len(self._workbook._font_styles)}">\n'
        for font_idx in sorted(self._workbook._font_styles.keys()):
            font_data = self._workbook._font_styles[font_idx]
            content += self._format_font_xml(font_data)
        content += '    </fonts>\n'
        
        # Write fills
        content += f'    <fills count="{len(self._workbook._fill_styles)}">\n'
        for fill_idx in sorted(self._workbook._fill_styles.keys()):
            fill_data = self._workbook._fill_styles[fill_idx]
            content += self._format_fill_xml(fill_data)
        content += '    </fills>\n'
        
        # Write borders
        content += f'    <borders count="{len(self._workbook._border_styles)}">\n'
        for border_idx in sorted(self._workbook._border_styles.keys()):
            border_data = self._workbook._border_styles[border_idx]
            content += self._format_border_xml(border_data)
        content += '    </borders>\n'
        
        # Write cellXfs
        content += f'    <cellXfs count="{len(self._workbook._cell_styles) + 1}">\n'
        # Default cellXf
        content += '        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'
        
        for cell_style_key, xf_idx in sorted(self._workbook._cell_styles.items(), key=lambda x: x[1]):
            font_idx, fill_idx, border_idx, num_fmt_idx, alignment_idx, protection_idx = cell_style_key
            apply_number_format = f' applyNumberFormat="1"' if num_fmt_idx != 0 else ''
            apply_protection = f' applyProtection="1"' if protection_idx != 0 else ''
            content += f'        <xf numFmtId="{num_fmt_idx}" fontId="{font_idx}" fillId="{fill_idx}" borderId="{border_idx}" xfId="0"{apply_number_format}{apply_protection}>\n'
            if alignment_idx > 0:
                align_data = self._workbook._alignment_styles[alignment_idx]
                content += self._format_alignment_xml(align_data)
            if protection_idx > 0:
                prot_data = self._workbook._protection_styles[protection_idx]
                content += self._format_protection_xml(prot_data)
            content += '        </xf>\n'
        
        content += '    </cellXfs>\n'

        # Write differential formatting (dxf) for conditional formatting
        if len(self._dxf_styles) > 0:
            content += f'    <dxfs count="{len(self._dxf_styles)}">\n'
            for dxf_data in self._dxf_styles:
                content += self._format_dxf_xml(dxf_data)
            content += '    </dxfs>\n'
        else:
            content += '    <dxfs count="0"/>\n'

        content += '</styleSheet>\n'
        zipf.writestr('xl/styles.xml', content)

    def _format_dxf_xml(self, dxf_data):
        """
        Formats differential formatting (dxf) as XML for conditional formatting.

        Args:
            dxf_data (dict): Dictionary containing font, fill, and border data.

        Returns:
            str: XML representation of the dxf element.
        """
        xml = '        <dxf>\n'

        # Add font if present
        if 'font' in dxf_data:
            font = dxf_data['font']
            xml += '            <font>\n'
            if font.get('bold'):
                xml += '                <b val="1"/>\n'
            if font.get('italic'):
                xml += '                <i val="1"/>\n'
            if font.get('underline'):
                xml += '                <u/>\n'
            if font.get('strikethrough'):
                xml += '                <strike/>\n'
            if font.get('color'):
                xml += f'                <color rgb="{font["color"]}"/>\n'
            xml += '            </font>\n'

        # Add fill if present
        if 'fill' in dxf_data:
            fill = dxf_data['fill']
            xml += '            <fill>\n'
            xml += f'                <patternFill patternType="{fill.get("pattern_type", "solid")}">\n'
            if fill.get('fg_color'):
                xml += f'                    <fgColor rgb="{fill["fg_color"]}"/>\n'
            if fill.get('bg_color'):
                xml += f'                    <bgColor rgb="{fill["bg_color"]}"/>\n'
            xml += '                </patternFill>\n'
            xml += '            </fill>\n'

        # Add border if present
        if 'border' in dxf_data:
            border = dxf_data['border']
            style = border.get('style', 'thin')
            color = border.get('color', 'FF000000')
            xml += '            <border>\n'
            xml += f'                <left style="{style}"><color rgb="{color}"/></left>\n'
            xml += f'                <right style="{style}"><color rgb="{color}"/></right>\n'
            xml += f'                <top style="{style}"><color rgb="{color}"/></top>\n'
            xml += f'                <bottom style="{style}"><color rgb="{color}"/></bottom>\n'
            xml += '            </border>\n'

        xml += '        </dxf>\n'
        return xml

    def _format_font_xml(self, font_data):
        """Formats font data as XML."""
        xml = '        <font>\n'
        if font_data.get('bold'):
            xml += '            <b/>\n'
        if font_data.get('italic'):
            xml += '            <i/>\n'
        if font_data.get('underline'):
            xml += '            <u/>\n'
        if font_data.get('strikethrough'):
            xml += '            <strike/>\n'
        xml += f'            <sz val="{font_data["size"]}"/>\n'
        xml += f'            <color rgb="{font_data["color"]}"/>\n'
        xml += f'            <name val="{font_data["name"]}"/>\n'
        xml += '        </font>\n'
        return xml
    
    def _format_fill_xml(self, fill_data):
        """Formats fill data as XML."""
        pattern_type = fill_data['pattern_type']
        xml = f'        <fill>\n            <patternFill patternType="{pattern_type}"'
        if pattern_type != 'none' and pattern_type != 'gray125':
            xml += f'>\n                <fgColor rgb="{fill_data["fg_color"]}"/>\n                <bgColor rgb="{fill_data["bg_color"]}"/>\n            </patternFill>\n'
        else:
            xml += '/>\n'
        xml += '        </fill>\n'
        return xml
    
    def _format_border_xml(self, border_data):
        """Formats border data as XML."""
        xml = '        <border>\n'
        for side in ['left', 'right', 'top', 'bottom']:
            side_data = border_data[side]
            if side_data['style'] != 'none':
                xml += f'            <{side} style="{side_data["style"]}">\n'
                xml += f'                <color rgb="{side_data["color"]}"/>\n'
                xml += f'            </{side}>\n'
        xml += '        </border>\n'
        return xml
    
    def _format_alignment_xml(self, align_data):
        """Formats alignment data as XML."""
        attrs = []
        if align_data['horizontal'] != 'general':
            attrs.append(f'horizontal="{align_data["horizontal"]}"')
        if align_data['vertical'] != 'bottom':
            attrs.append(f'vertical="{align_data["vertical"]}"')
        if align_data['wrap_text']:
            attrs.append('wrapText="1"')
        if align_data['indent'] != 0:
            attrs.append(f'indent="{align_data["indent"]}"')
        if align_data['text_rotation'] != 0:
            attrs.append(f'textRotation="{align_data["text_rotation"]}"')
        if align_data['shrink_to_fit']:
            attrs.append('shrinkToFit="1"')
        if align_data['reading_order'] != 0:
            attrs.append(f'readingOrder="{align_data["reading_order"]}"')
        if align_data['relative_indent'] != 0:
            attrs.append(f'relativeIndent="{align_data["relative_indent"]}"')
        
        if attrs:
            xml = '            <alignment ' + ' '.join(attrs) + '/>\n'
        else:
            xml = '            <alignment/>\n'
        return xml

    def _format_protection_xml(self, prot_data):
        """
        Formats protection data as XML.

        ECMA-376 Section: 18.8.33 (protection element)
        Default values: locked="1", hidden="0"
        """
        attrs = []
        # Only write non-default values
        if not prot_data['locked']:  # Default is True (1)
            attrs.append('locked="0"')
        if prot_data['hidden']:  # Default is False (0)
            attrs.append('hidden="1"')

        if attrs:
            xml = '            <protection ' + ' '.join(attrs) + '/>\n'
        else:
            # If both are default, we can omit the element entirely
            xml = ''
        return xml

    def _write_shared_strings_xml(self, zipf):
        """Writes xl/sharedStrings.xml file."""
        content = self._shared_string_table.to_xml()
        zipf.writestr('xl/sharedStrings.xml', content)
    
    # Style management methods for XML creation
    
    def register_default_styles(self):
        """Registers default styles for fonts, fills, borders, and alignments."""
        # Default font (Calibri, 11pt, black) - index 0
        self._workbook._font_styles[0] = {
            'name': 'Calibri',
            'size': 11,
            'color': 'FF000000',
            'bold': False,
            'italic': False,
            'underline': False,
            'strikethrough': False
        }
        
        # Default fills
        self._workbook._fill_styles[0] = {  # No fill
            'pattern_type': 'none',
            'fg_color': 'FFFFFFFF',
            'bg_color': 'FFFFFFFF'
        }
        self._workbook._fill_styles[1] = {  # Gray pattern
            'pattern_type': 'gray125',
            'fg_color': 'FFFFFFFF',
            'bg_color': 'FFFFFFFF'
        }
        
        # Default borders
        self._workbook._border_styles[0] = {
            'top': {'style': 'none', 'color': 'FF000000'},
            'bottom': {'style': 'none', 'color': 'FF000000'},
            'left': {'style': 'none', 'color': 'FF000000'},
            'right': {'style': 'none', 'color': 'FF000000'}
        }

        # Default protection (locked=True, hidden=False)
        self._workbook._protection_styles[0] = {
            'locked': True,
            'hidden': False
        }
        
        # Default alignment (general/bottom) - index 0
        self._workbook._alignment_styles[0] = {
            'horizontal': 'general',
            'vertical': 'bottom',
            'wrap_text': False,
            'indent': 0,
            'text_rotation': 0,
            'shrink_to_fit': False,
            'reading_order': 0,
            'relative_indent': 0
        }
    
    def get_or_create_font_style(self, font):
        """Gets or creates a font style index."""
        # Check if this font already exists by comparing with existing fonts
        for idx, font_data in self._workbook._font_styles.items():
            if (font_data['name'] == font.name and
                font_data['size'] == font.size and
                font_data['color'] == font.color and
                font_data['bold'] == font.bold and
                font_data['italic'] == font.italic and
                font_data['underline'] == font.underline and
                font_data['strikethrough'] == font.strikethrough):
                return idx
        
        # Create new font style
        new_idx = len(self._workbook._font_styles)
        self._workbook._font_styles[new_idx] = {
            'name': font.name,
            'size': font.size,
            'color': font.color,
            'bold': font.bold,
            'italic': font.italic,
            'underline': font.underline,
            'strikethrough': font.strikethrough
        }
        return new_idx
    
    def get_or_create_fill_style(self, fill):
        """Gets or creates a fill style index."""
        # Check if this fill already exists by comparing with existing fills
        for idx, fill_data in self._workbook._fill_styles.items():
            if (fill_data['pattern_type'] == fill.pattern_type and
                fill_data['fg_color'] == fill.foreground_color and
                fill_data['bg_color'] == fill.background_color):
                return idx
        
        # Create new fill style
        new_idx = len(self._workbook._fill_styles)
        self._workbook._fill_styles[new_idx] = {
            'pattern_type': fill.pattern_type,
            'fg_color': fill.foreground_color,
            'bg_color': fill.background_color
        }
        return new_idx
    
    def get_or_create_border_style(self, borders):
        """Gets or creates a border style index."""
        # Check if this border already exists by comparing with existing borders
        for idx, border_data in self._workbook._border_styles.items():
            if (border_data['top']['style'] == borders.top.line_style and
                border_data['top']['color'] == borders.top.color and
                border_data['bottom']['style'] == borders.bottom.line_style and
                border_data['bottom']['color'] == borders.bottom.color and
                border_data['left']['style'] == borders.left.line_style and
                border_data['left']['color'] == borders.left.color and
                border_data['right']['style'] == borders.right.line_style and
                border_data['right']['color'] == borders.right.color):
                return idx
        
        # Create new border style with all four sides
        new_idx = len(self._workbook._border_styles)
        self._workbook._border_styles[new_idx] = {
            'top': {'style': borders.top.line_style, 'color': borders.top.color},
            'bottom': {'style': borders.bottom.line_style, 'color': borders.bottom.color},
            'left': {'style': borders.left.line_style, 'color': borders.left.color},
            'right': {'style': borders.right.line_style, 'color': borders.right.color}
        }
        return new_idx
    
    def get_or_create_alignment_style(self, alignment):
        """Gets or creates an alignment style index."""
        # Check if this alignment already exists by comparing with existing alignments
        for idx, align_data in self._workbook._alignment_styles.items():
            if (align_data['horizontal'] == alignment.horizontal and
                align_data['vertical'] == alignment.vertical and
                align_data['wrap_text'] == alignment.wrap_text and
                align_data['indent'] == alignment.indent and
                align_data['text_rotation'] == alignment.text_rotation and
                align_data['shrink_to_fit'] == alignment.shrink_to_fit and
                align_data['reading_order'] == alignment.reading_order and
                align_data['relative_indent'] == alignment.relative_indent):
                return idx
        
        # Create new alignment style
        new_idx = len(self._workbook._alignment_styles)
        self._workbook._alignment_styles[new_idx] = {
            'horizontal': alignment.horizontal,
            'vertical': alignment.vertical,
            'wrap_text': alignment.wrap_text,
            'indent': alignment.indent,
            'text_rotation': alignment.text_rotation,
            'shrink_to_fit': alignment.shrink_to_fit,
            'reading_order': alignment.reading_order,
            'relative_indent': alignment.relative_indent
        }
        return new_idx

    def get_or_create_protection_style(self, protection):
        """Gets or creates a protection style index."""
        # Check if this protection already exists by comparing with existing protection styles
        for idx, prot_data in self._workbook._protection_styles.items():
            if (prot_data['locked'] == protection.locked and
                prot_data['hidden'] == protection.hidden):
                return idx

        # Create new protection style
        new_idx = len(self._workbook._protection_styles)
        self._workbook._protection_styles[new_idx] = {
            'locked': protection.locked,
            'hidden': protection.hidden
        }
        return new_idx

    def get_or_create_number_format_style(self, number_format):
        """Gets or creates a number format style index."""
        # Built-in number formats (0-163)
        builtin_formats = {
            'General': 0,
            '0': 1,
            '0.00': 2,
            '#,##0': 3,
            '#,##0.00': 4,
            '$#,##0_);($#,##0)': 5,
            '$#,##0_);[Red]($#,##0)': 6,
            '$#,##0.00_);($#,##0.00)': 7,
            '$#,##0.00_);[Red]($#,##0.00)': 8,
            '0%': 9,
            '0.00%': 10,
            '0.00E+00': 11,
            '# ?/?': 12,
            '# ??/??': 13,
            'mm-dd-yy': 14,
            'd-mmm-yy': 15,
            'd-mmm': 16,
            'mmm-yy': 17,
            'h:mm AM/PM': 18,
            'h:mm:ss AM/PM': 19,
            'h:mm': 20,
            'h:mm:ss': 21,
            'm/d/yy h:mm': 22,
            '#,##0_);(#,##0)': 37,
            '#,##0_);[Red](#,##0)': 38,
            '#,##0.00_);(#,##0.00)': 39,
            '#,##0.00_);[Red](#,##0.00)': 40,
            'mm:ss': 45,
            '[h]:mm:ss': 46,
            'mm:ss.0': 47,
            '##0.0E+0': 48,
            '@': 49
        }
        
        # Check if this is a built-in format
        if number_format in builtin_formats:
            return builtin_formats[number_format]
        
        # Check if this custom number format already exists
        for idx, fmt in self._workbook._num_formats.items():
            if fmt == number_format:
                return idx
        
        # Create new custom number format style (start from ID 164)
        new_idx = 164 + len([k for k in self._workbook._num_formats.keys() if k >= 164])
        self._workbook._num_formats[new_idx] = number_format
        return new_idx
    
    def get_or_create_cell_style(self, cell):
        """Gets or creates a cell xf style index."""
        font_idx = self.get_or_create_font_style(cell.style.font)
        fill_idx = self.get_or_create_fill_style(cell.style.fill)
        border_idx = self.get_or_create_border_style(cell.style.borders)
        num_fmt_idx = self.get_or_create_number_format_style(cell.style.number_format)
        alignment_idx = self.get_or_create_alignment_style(cell.style.alignment)
        protection_idx = self.get_or_create_protection_style(cell.style.protection)

        # Check if this is the default style (all indices are 0)
        # If so, return 0 to use the default xf in cellXfs
        if (font_idx == 0 and fill_idx == 0 and border_idx == 0 and
            num_fmt_idx == 0 and alignment_idx == 0 and protection_idx == 0):
            return 0

        key = (font_idx, fill_idx, border_idx, num_fmt_idx, alignment_idx, protection_idx)
        if key not in self._workbook._cell_styles:
            # The cellXfs collection has a default xf at index 0, so custom xfs start at index 1
            # We store the actual xf index (starting from 1) in _cell_styles dictionary
            self._workbook._cell_styles[key] = len(self._workbook._cell_styles) + 1
        return self._workbook._cell_styles[key]

    def _write_core_properties_xml(self, zipf):
        """
        Writes docProps/core.xml file with core document properties.
        
        ECMA-376 Part 2, Section 11 - Core Properties
        
        Uses Dublin Core (dc:) and OPC Core Properties (cp:) namespaces.
        """
        from datetime import datetime, timezone
        
        # Get document properties if available
        doc_props = getattr(self._workbook, 'document_properties', None)

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        content += 'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        content += 'xmlns:dcterms="http://purl.org/dc/terms/" '
        content += 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        content += 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'

        if doc_props and doc_props.core:
            core = doc_props.core

            if core.title:
                content += f'    <dc:title>{self._escape_xml(core.title)}</dc:title>\n'
            if core.subject:
                content += f'    <dc:subject>{self._escape_xml(core.subject)}</dc:subject>\n'
            if core.creator:
                content += f'    <dc:creator>{self._escape_xml(core.creator)}</dc:creator>\n'
            if core.keywords:
                content += f'    <cp:keywords>{self._escape_xml(core.keywords)}</cp:keywords>\n'
            if core.description:
                content += f'    <dc:description>{self._escape_xml(core.description)}</dc:description>\n'
            if core.last_modified_by:
                content += f'    <cp:lastModifiedBy>{self._escape_xml(core.last_modified_by)}</cp:lastModifiedBy>\n'
            if core.revision:
                content += f'    <cp:revision>{self._escape_xml(str(core.revision))}</cp:revision>\n'
            if core.category:
                content += f'    <cp:category>{self._escape_xml(core.category)}</cp:category>\n'
            if core.content_status:
                content += f'    <cp:contentStatus>{self._escape_xml(core.content_status)}</cp:contentStatus>\n'

            # Handle dates - format as W3CDTF
            if core.created:
                if isinstance(core.created, datetime):
                    created_str = core.created.strftime('%Y-%m-%dT%H:%M:%SZ')
                else:
                    created_str = str(core.created)
                content += f'    <dcterms:created xsi:type="dcterms:W3CDTF">{created_str}</dcterms:created>\n'

            if core.modified:
                if isinstance(core.modified, datetime):
                    modified_str = core.modified.strftime('%Y-%m-%dT%H:%M:%SZ')
                else:
                    modified_str = str(core.modified)
                content += f'    <dcterms:modified xsi:type="dcterms:W3CDTF">{modified_str}</dcterms:modified>\n'
        else:
            # Write default created/modified dates
            now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
            content += f'    <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>\n'
            content += f'    <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>\n'

        content += '</cp:coreProperties>\n'
        zipf.writestr('docProps/core.xml', content)

    def _write_app_properties_xml(self, zipf):
        """
        Writes docProps/app.xml file with extended/application properties.

        ECMA-376 Part 1, Section 22.2 - Extended Properties
        """
        # Get document properties if available
        doc_props = getattr(self._workbook, 'document_properties', None)

        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        content += 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n'

        if doc_props and doc_props.extended:
            ext = doc_props.extended

            content += f'    <Application>{self._escape_xml(ext.application or "Microsoft Excel")}</Application>\n'
            content += f'    <DocSecurity>{ext.doc_security}</DocSecurity>\n'
            content += f'    <ScaleCrop>{"true" if ext.scale_crop else "false"}</ScaleCrop>\n'

            if ext.company:
                content += f'    <Company>{self._escape_xml(ext.company)}</Company>\n'
            if ext.manager:
                content += f'    <Manager>{self._escape_xml(ext.manager)}</Manager>\n'
            if ext.hyperlink_base:
                content += f'    <HyperlinkBase>{self._escape_xml(ext.hyperlink_base)}</HyperlinkBase>\n'
            if ext.app_version:
                content += f'    <AppVersion>{self._escape_xml(ext.app_version)}</AppVersion>\n'

            content += f'    <LinksUpToDate>{"true" if ext.links_up_to_date else "false"}</LinksUpToDate>\n'
            content += f'    <SharedDoc>{"true" if ext.shared_doc else "false"}</SharedDoc>\n'
        else:
            # Write default values
            content += '    <Application>Microsoft Excel</Application>\n'
            content += '    <DocSecurity>0</DocSecurity>\n'
            content += '    <ScaleCrop>false</ScaleCrop>\n'
            content += '    <LinksUpToDate>false</LinksUpToDate>\n'
            content += '    <SharedDoc>false</SharedDoc>\n'

        # Add heading pairs and titles of parts (worksheet names)
        worksheet_count = len(self._workbook.worksheets)
        content += '    <HeadingPairs>\n'
        content += '        <vt:vector size="2" baseType="variant">\n'
        content += '            <vt:variant>\n'
        content += '                <vt:lpstr>Worksheets</vt:lpstr>\n'
        content += '            </vt:variant>\n'
        content += '            <vt:variant>\n'
        content += f'                <vt:i4>{worksheet_count}</vt:i4>\n'
        content += '            </vt:variant>\n'
        content += '        </vt:vector>\n'
        content += '    </HeadingPairs>\n'

        content += '    <TitlesOfParts>\n'
        content += f'        <vt:vector size="{worksheet_count}" baseType="lpstr">\n'
        for worksheet in self._workbook.worksheets:
            content += f'            <vt:lpstr>{self._escape_xml(worksheet.name)}</vt:lpstr>\n'
        content += '        </vt:vector>\n'
        content += '    </TitlesOfParts>\n'

        content += '</Properties>\n'
        zipf.writestr('docProps/app.xml', content)
