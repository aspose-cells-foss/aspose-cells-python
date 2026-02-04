"""
Aspose.Cells for Python - XML Properties Saver Module

This module provides classes for saving workbook and worksheet properties
to XML format according to ECMA-376 specification.

ECMA-376 Sections: 18.2 (Workbook), 18.3.1 (Worksheet)
"""


class WorkbookPropertiesXMLWriter:
    """
    Handles writing workbook properties to XML format for .xlsx files.

    Examples:
        >>> writer = WorkbookPropertiesXMLWriter(escape_xml_func)
        >>> xml = writer.format_file_version_xml(workbook.properties.file_version)
    """

    def __init__(self, escape_xml_func):
        """
        Initializes a new instance of the WorkbookPropertiesXMLWriter class.

        Args:
            escape_xml_func: A function to escape XML special characters.
        """
        self._escape_xml = escape_xml_func

    def format_file_version_xml(self, file_version):
        """
        Formats file version element as XML.

        ECMA-376 Section: 18.2.10

        Args:
            file_version: FileVersion object.

        Returns:
            str: XML representation of the file version.
        """
        attrs = []
        if file_version.app_name:
            attrs.append(f'appName="{self._escape_xml(file_version.app_name)}"')
        if file_version.last_edited is not None:
            attrs.append(f'lastEdited="{file_version.last_edited}"')
        if file_version.lowest_edited is not None:
            attrs.append(f'lowestEdited="{file_version.lowest_edited}"')
        if file_version.rup_build is not None:
            attrs.append(f'rupBuild="{file_version.rup_build}"')

        if attrs:
            return f'    <fileVersion {" ".join(attrs)}/>\n'
        return ''

    def format_workbook_protection_xml(self, protection):
        """
        Formats workbook protection element as XML.

        ECMA-376 Section: 18.2.29

        Args:
            protection: WorkbookProtection object.

        Returns:
            str: XML representation of workbook protection.
        """
        if not protection.is_protected:
            return ''

        attrs = []
        if protection.lock_structure:
            attrs.append('lockStructure="1"')
        if protection.lock_windows:
            attrs.append('lockWindows="1"')
        if protection.lock_revision:
            attrs.append('lockRevision="1"')
        if protection.workbook_password:
            attrs.append(f'workbookPassword="{self._escape_xml(protection.workbook_password)}"')
        if protection.revisions_password:
            attrs.append(f'revisionsPassword="{self._escape_xml(protection.revisions_password)}"')

        if attrs:
            return f'    <workbookProtection {" ".join(attrs)}/>\n'
        return ''

    def format_workbook_pr_xml(self, workbook_pr):
        """
        Formats workbook properties (workbookPr) element as XML.

        ECMA-376 Section: 18.2.13

        Args:
            workbook_pr: WorkbookPr object.

        Returns:
            str: XML representation of workbookPr.
        """
        attrs = []
        if workbook_pr.date1904:
            attrs.append('date1904="1"')
        if workbook_pr.code_name:
            attrs.append(f'codeName="{self._escape_xml(workbook_pr.code_name)}"')
        if workbook_pr.show_objects != "all":
            attrs.append(f'showObjects="{workbook_pr.show_objects}"')
        if workbook_pr.filter_privacy:
            attrs.append('filterPrivacy="1"')
        if not workbook_pr.show_border_unselected_tables:
            attrs.append('showBorderUnselectedTables="0"')
        if not workbook_pr.show_ink_annotation:
            attrs.append('showInkAnnotation="0"')
        if workbook_pr.backup_file:
            attrs.append('backupFile="1"')
        if not workbook_pr.save_external_link_values:
            attrs.append('saveExternalLinkValues="0"')
        if workbook_pr.update_links != "userSet":
            attrs.append(f'updateLinks="{workbook_pr.update_links}"')
        if workbook_pr.hide_pivot_field_list:
            attrs.append('hidePivotFieldList="1"')
        if workbook_pr.default_theme_version is not None:
            attrs.append(f'defaultThemeVersion="{workbook_pr.default_theme_version}"')

        if attrs:
            return f'    <workbookPr {" ".join(attrs)}/>\n'
        return ''

    def format_book_views_xml(self, view):
        """
        Formats book views element as XML.

        ECMA-376 Section: 18.2.1, 18.2.30

        Args:
            view: WorkbookView object.

        Returns:
            str: XML representation of book views.
        """
        attrs = []
        attrs.append(f'xWindow="{view.x_window}"')
        attrs.append(f'yWindow="{view.y_window}"')
        attrs.append(f'windowWidth="{view.window_width}"')
        attrs.append(f'windowHeight="{view.window_height}"')

        if view.active_tab != 0:
            attrs.append(f'activeTab="{view.active_tab}"')
        if view.first_sheet != 0:
            attrs.append(f'firstSheet="{view.first_sheet}"')
        if not view.show_horizontal_scroll:
            attrs.append('showHorizontalScroll="0"')
        if not view.show_vertical_scroll:
            attrs.append('showVerticalScroll="0"')
        if not view.show_sheet_tabs:
            attrs.append('showSheetTabs="0"')
        if view.tab_ratio != 600:
            attrs.append(f'tabRatio="{view.tab_ratio}"')
        if view.visibility != 'visible':
            attrs.append(f'visibility="{view.visibility}"')
        if view.minimized:
            attrs.append('minimized="1"')

        xml = '    <bookViews>\n'
        xml += f'        <workbookView {" ".join(attrs)}/>\n'
        xml += '    </bookViews>\n'
        return xml

    def format_calc_pr_xml(self, calculation):
        """
        Formats calculation properties element as XML.

        ECMA-376 Section: 18.2.2

        Args:
            calculation: CalculationProperties object.

        Returns:
            str: XML representation of calculation properties.
        """
        attrs = []
        if calculation.calc_id is not None:
            attrs.append(f'calcId="{calculation.calc_id}"')
        if calculation.calc_mode != 'auto':
            attrs.append(f'calcMode="{calculation.calc_mode}"')
        if calculation.full_calc_on_load:
            attrs.append('fullCalcOnLoad="1"')
        if calculation.ref_mode != 'A1':
            attrs.append(f'refMode="{calculation.ref_mode}"')
        if calculation.iterate:
            attrs.append('iterate="1"')
            if calculation.iterate_count != 100:
                attrs.append(f'iterateCount="{calculation.iterate_count}"')
            if calculation.iterate_delta != 0.001:
                attrs.append(f'iterateDelta="{calculation.iterate_delta}"')
        if not calculation.full_precision:
            attrs.append('fullPrecision="0"')
        if not calculation.calc_on_save:
            attrs.append('calcOnSave="0"')
        if not calculation.concurrent_calc:
            attrs.append('concurrentCalc="0"')
        if not calculation.calc_completed:
            attrs.append('calcCompleted="0"')
        if calculation.force_full_calc:
            attrs.append('forceFullCalc="1"')

        if attrs:
            return f'    <calcPr {" ".join(attrs)}/>\n'
        return ''

    def format_defined_names_xml(self, defined_names):
        """
        Formats defined names element as XML.

        ECMA-376 Section: 18.2.6

        Args:
            defined_names: DefinedNameCollection object.

        Returns:
            str: XML representation of defined names.
        """
        if len(defined_names) == 0:
            return ''

        xml = '    <definedNames>\n'
        for dn in defined_names:
            attrs = [f'name="{self._escape_xml(dn.name)}"']
            if dn.local_sheet_id is not None:
                attrs.append(f'localSheetId="{dn.local_sheet_id}"')
            if dn.hidden:
                attrs.append('hidden="1"')
            if dn.comment:
                attrs.append(f'comment="{self._escape_xml(dn.comment)}"')

            refers_to = self._escape_xml(dn.refers_to)
            xml += f'        <definedName {" ".join(attrs)}>{refers_to}</definedName>\n'
        xml += '    </definedNames>\n'
        return xml


class WorksheetPropertiesXMLWriter:
    """
    Handles writing worksheet properties to XML format for .xlsx files.

    Examples:
        >>> writer = WorksheetPropertiesXMLWriter(escape_xml_func)
        >>> xml = writer.format_sheet_views_xml(worksheet.properties)
    """

    def __init__(self, escape_xml_func):
        """
        Initializes a new instance of the WorksheetPropertiesXMLWriter class.

        Args:
            escape_xml_func: A function to escape XML special characters.
        """
        self._escape_xml = escape_xml_func

    def format_sheet_views_xml(self, properties, is_selected=False):
        """
        Formats sheet views element as XML.

        ECMA-376 Section: 18.3.1.88, 18.3.1.87

        Args:
            properties: WorksheetProperties object.
            is_selected: Whether this sheet is the selected/active sheet.

        Returns:
            str: XML representation of sheet views.
        """
        view = properties.view
        selection = properties.selection
        pane = properties.pane

        attrs = []
        if view.show_formulas:
            attrs.append('showFormulas="1"')
        if not view.show_grid_lines:
            attrs.append('showGridLines="0"')
        if not view.show_row_col_headers:
            attrs.append('showRowColHeaders="0"')
        if not view.show_zeros:
            attrs.append('showZeros="0"')
        if view.right_to_left:
            attrs.append('rightToLeft="1"')
        if is_selected or view.tab_selected:
            attrs.append('tabSelected="1"')
        if not view.show_ruler:
            attrs.append('showRuler="0"')
        if not view.show_outline_symbols:
            attrs.append('showOutlineSymbols="0"')
        if not view.default_grid_color:
            attrs.append('defaultGridColor="0"')
        if not view.show_white_space:
            attrs.append('showWhiteSpace="0"')
        if view.view != 'normal':
            attrs.append(f'view="{view.view}"')
        if view.top_left_cell:
            attrs.append(f'topLeftCell="{view.top_left_cell}"')
        if view.color_id is not None and view.color_id != 64:
            attrs.append(f'colorId="{view.color_id}"')
        if view.zoom_scale != 100:
            attrs.append(f'zoomScale="{view.zoom_scale}"')
        if view.zoom_scale_normal != 0:
            attrs.append(f'zoomScaleNormal="{view.zoom_scale_normal}"')
        if view.zoom_scale_sheet_layout_view != 0:
            attrs.append(f'zoomScaleSheetLayoutView="{view.zoom_scale_sheet_layout_view}"')
        if view.zoom_scale_page_layout_view != 0:
            attrs.append(f'zoomScalePageLayoutView="{view.zoom_scale_page_layout_view}"')
        if view.window_protection:
            attrs.append('windowProtection="1"')
        attrs.append(f'workbookViewId="{view.workbook_view_id}"')

        xml = '    <sheetViews>\n'
        xml += f'        <sheetView {" ".join(attrs)}'

        # Check if we need child elements (pane or selection)
        has_pane = pane.state is not None
        has_selection = selection.active_cell != 'A1' or selection.sqref != 'A1'

        if has_pane or has_selection:
            xml += '>\n'

            # Add pane element if frozen/split
            if has_pane:
                xml += self._format_pane_xml(pane)

            # Add selection element
            sel_attrs = []
            if pane.state and pane.active_pane:
                sel_attrs.append(f'pane="{pane.active_pane}"')
            sel_attrs.append(f'activeCell="{selection.active_cell}"')
            sel_attrs.append(f'sqref="{selection.sqref}"')
            xml += f'            <selection {" ".join(sel_attrs)}/>\n'

            xml += '        </sheetView>\n'
        else:
            xml += '/>\n'

        xml += '    </sheetViews>\n'
        return xml

    def _format_pane_xml(self, pane):
        """Formats pane element as XML."""
        attrs = []
        if pane.x_split is not None and pane.x_split > 0:
            attrs.append(f'xSplit="{pane.x_split}"')
        if pane.y_split is not None and pane.y_split > 0:
            attrs.append(f'ySplit="{pane.y_split}"')
        if pane.top_left_cell:
            attrs.append(f'topLeftCell="{pane.top_left_cell}"')
        if pane.active_pane:
            attrs.append(f'activePane="{pane.active_pane}"')
        if pane.state:
            attrs.append(f'state="{pane.state}"')

        return f'            <pane {" ".join(attrs)}/>\n'

    def format_sheet_format_pr_xml(self, format_props):
        """
        Formats sheet format properties element as XML.

        ECMA-376 Section: 18.3.1.82

        Args:
            format_props: SheetFormatProperties object.

        Returns:
            str: XML representation of sheet format properties.
        """
        attrs = []
        if format_props.base_col_width != 8:
            attrs.append(f'baseColWidth="{format_props.base_col_width}"')
        if format_props.default_col_width is not None:
            attrs.append(f'defaultColWidth="{format_props.default_col_width}"')
        attrs.append(f'defaultRowHeight="{format_props.default_row_height}"')
        if format_props.custom_height:
            attrs.append('customHeight="1"')
        if format_props.zero_height:
            attrs.append('zeroHeight="1"')
        if format_props.thick_top:
            attrs.append('thickTop="1"')
        if format_props.thick_bottom:
            attrs.append('thickBottom="1"')
        if format_props.outline_level_row > 0:
            attrs.append(f'outlineLevelRow="{format_props.outline_level_row}"')
        if format_props.outline_level_col > 0:
            attrs.append(f'outlineLevelCol="{format_props.outline_level_col}"')

        return f'    <sheetFormatPr {" ".join(attrs)}/>\n'

    def format_sheet_protection_xml(self, protection):
        """
        Formats sheet protection element as XML.

        ECMA-376 Section: 18.3.1.85

        Args:
            protection: SheetProtection object.

        Returns:
            str: XML representation of sheet protection.
        """
        if not protection.is_protected:
            return ''

        from .workbook_hash_password import hash_password

        attrs = []
        if protection.sheet:
            attrs.append('sheet="1"')
        if protection.objects:
            attrs.append('objects="1"')
        if protection.scenarios:
            attrs.append('scenarios="1"')
        if not protection.format_cells:
            attrs.append('formatCells="0"')
        if not protection.format_columns:
            attrs.append('formatColumns="0"')
        if not protection.format_rows:
            attrs.append('formatRows="0"')
        if not protection.insert_columns:
            attrs.append('insertColumns="0"')
        if not protection.insert_rows:
            attrs.append('insertRows="0"')
        if not protection.insert_hyperlinks:
            attrs.append('insertHyperlinks="0"')
        if not protection.delete_columns:
            attrs.append('deleteColumns="0"')
        if not protection.delete_rows:
            attrs.append('deleteRows="0"')
        if protection.select_locked_cells:
            attrs.append('selectLockedCells="1"')
        if not protection.sort:
            attrs.append('sort="0"')
        if not protection.auto_filter:
            attrs.append('autoFilter="0"')
        if not protection.pivot_tables:
            attrs.append('pivotTables="0"')
        if protection.select_unlocked_cells:
            attrs.append('selectUnlockedCells="1"')
        if protection.password:
            # Hash the password for Excel compatibility (if not already hashed)
            password_value = protection.password
            # Check if password looks like a hash (4 hex digits)
            if not (len(password_value) == 4 and all(c in '0123456789ABCDEFabcdef' for c in password_value)):
                # Password is plaintext, hash it
                password_value = hash_password(password_value)
            attrs.append(f'password="{self._escape_xml(password_value)}"')
        if protection.algorithm_name:
            attrs.append(f'algorithmName="{self._escape_xml(protection.algorithm_name)}"')
        if protection.hash_value:
            attrs.append(f'hashValue="{self._escape_xml(protection.hash_value)}"')
        if protection.salt_value:
            attrs.append(f'saltValue="{self._escape_xml(protection.salt_value)}"')
        if protection.spin_count is not None:
            attrs.append(f'spinCount="{protection.spin_count}"')

        if attrs:
            return f'    <sheetProtection {" ".join(attrs)}/>\n'
        return ''

    def format_print_options_xml(self, print_options):
        """
        Formats print options element as XML.

        ECMA-376 Section: 18.3.1.70

        Args:
            print_options: PrintOptions object.

        Returns:
            str: XML representation of print options.
        """
        attrs = []
        if print_options.print_headings:
            attrs.append('headings="1"')
        if print_options.print_grid_lines:
            attrs.append('gridLines="1"')
        if print_options.horizontal_centered:
            attrs.append('horizontalCentered="1"')
        if print_options.vertical_centered:
            attrs.append('verticalCentered="1"')
        if print_options.black_and_white:
            attrs.append('blackAndWhite="1"')
        if print_options.draft_quality:
            attrs.append('draft="1"')
        if print_options.cell_comments != 'none':
            attrs.append(f'comments="{print_options.cell_comments}"')
        if print_options.print_errors != 'displayed':
            attrs.append(f'errors="{print_options.print_errors}"')

        if attrs:
            return f'    <printOptions {" ".join(attrs)}/>\n'
        return ''

    def format_page_margins_xml(self, margins):
        """
        Formats page margins element as XML.

        ECMA-376 Section: 18.3.1.62

        Args:
            margins: PageMargins object.

        Returns:
            str: XML representation of page margins.
        """
        return (f'    <pageMargins left="{margins.left}" right="{margins.right}" '
                f'top="{margins.top}" bottom="{margins.bottom}" '
                f'header="{margins.header}" footer="{margins.footer}"/>\n')

    def format_page_setup_xml(self, page_setup):
        """
        Formats page setup element as XML.

        ECMA-376 Section: 18.3.1.63

        Args:
            page_setup: PageSetup object.

        Returns:
            str: XML representation of page setup.
        """
        attrs = []
        if page_setup.paper_size != 1:
            attrs.append(f'paperSize="{page_setup.paper_size}"')
        if page_setup.scale != 100:
            attrs.append(f'scale="{page_setup.scale}"')
        if page_setup.first_page_number is not None:
            attrs.append(f'firstPageNumber="{page_setup.first_page_number}"')
        if page_setup.fit_to_width is not None:
            attrs.append(f'fitToWidth="{page_setup.fit_to_width}"')
        if page_setup.fit_to_height is not None:
            attrs.append(f'fitToHeight="{page_setup.fit_to_height}"')
        if page_setup.page_order != 'downThenOver':
            attrs.append(f'pageOrder="{page_setup.page_order}"')
        if page_setup.orientation != 'portrait':
            attrs.append(f'orientation="{page_setup.orientation}"')
        if not page_setup.use_printer_defaults:
            attrs.append('usePrinterDefaults="0"')
        if page_setup.black_and_white:
            attrs.append('blackAndWhite="1"')
        if page_setup.draft:
            attrs.append('draft="1"')
        if page_setup.cell_comments != 'none':
            attrs.append(f'comments="{page_setup.cell_comments}"')
        if page_setup.use_first_page_number:
            attrs.append('useFirstPageNumber="1"')
        if page_setup.errors != 'displayed':
            attrs.append(f'errors="{page_setup.errors}"')
        if page_setup.horizontal_dpi is not None:
            attrs.append(f'horizontalDpi="{page_setup.horizontal_dpi}"')
        if page_setup.vertical_dpi is not None:
            attrs.append(f'verticalDpi="{page_setup.vertical_dpi}"')
        if page_setup.copies != 1:
            attrs.append(f'copies="{page_setup.copies}"')

        if attrs:
            return f'    <pageSetup {" ".join(attrs)}/>\n'
        return ''

    def format_header_footer_xml(self, header_footer):
        """
        Formats header/footer element as XML.

        ECMA-376 Section: 18.3.1.46

        Args:
            header_footer: HeaderFooter object.

        Returns:
            str: XML representation of header/footer.
        """
        has_content = (header_footer.odd_header or header_footer.odd_footer or
                       header_footer.even_header or header_footer.even_footer or
                       header_footer.first_header or header_footer.first_footer)

        if not has_content:
            return ''

        attrs = []
        if header_footer.different_first:
            attrs.append('differentFirst="1"')
        if header_footer.different_odd_even:
            attrs.append('differentOddEven="1"')
        if not header_footer.scale_with_doc:
            attrs.append('scaleWithDoc="0"')
        if not header_footer.align_with_margins:
            attrs.append('alignWithMargins="0"')

        xml = f'    <headerFooter {" ".join(attrs)}>\n' if attrs else '    <headerFooter>\n'

        if header_footer.odd_header:
            xml += f'        <oddHeader>{self._escape_xml(header_footer.odd_header)}</oddHeader>\n'
        if header_footer.odd_footer:
            xml += f'        <oddFooter>{self._escape_xml(header_footer.odd_footer)}</oddFooter>\n'
        if header_footer.even_header:
            xml += f'        <evenHeader>{self._escape_xml(header_footer.even_header)}</evenHeader>\n'
        if header_footer.even_footer:
            xml += f'        <evenFooter>{self._escape_xml(header_footer.even_footer)}</evenFooter>\n'
        if header_footer.first_header:
            xml += f'        <firstHeader>{self._escape_xml(header_footer.first_header)}</firstHeader>\n'
        if header_footer.first_footer:
            xml += f'        <firstFooter>{self._escape_xml(header_footer.first_footer)}</firstFooter>\n'

        xml += '    </headerFooter>\n'
        return xml
