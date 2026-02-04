"""
Aspose.Cells for Python - XML Properties Loader Module

This module provides classes for loading workbook and worksheet properties
from XML format according to ECMA-376 specification.

ECMA-376 Sections: 18.2 (Workbook), 18.3.1 (Worksheet)
"""


class WorkbookPropertiesXMLLoader:
    """
    Handles loading workbook properties from XML format for .xlsx files.

    Examples:
        >>> loader = WorkbookPropertiesXMLLoader(namespaces)
        >>> loader.load_file_version(workbook.properties.file_version, root)
    """

    def __init__(self, namespaces):
        """
        Initializes a new instance of the WorkbookPropertiesXMLLoader class.

        Args:
            namespaces: XML namespaces dictionary for parsing.
        """
        self.ns = namespaces

    def load_file_version(self, file_version, root):
        """
        Loads file version from XML.

        ECMA-376 Section: 18.2.10

        Args:
            file_version: FileVersion object to load into.
            root: XML root element.
        """
        elem = root.find('main:fileVersion', namespaces=self.ns)
        if elem is None:
            return

        if elem.get('appName'):
            file_version._app_name = elem.get('appName')
        if elem.get('lastEdited'):
            file_version._last_edited = int(elem.get('lastEdited'))
        if elem.get('lowestEdited'):
            file_version._lowest_edited = int(elem.get('lowestEdited'))
        if elem.get('rupBuild'):
            file_version._rup_build = int(elem.get('rupBuild'))

    def load_workbook_protection(self, protection, root):
        """
        Loads workbook protection from XML.

        ECMA-376 Section: 18.2.29

        Args:
            protection: WorkbookProtection object to load into.
            root: XML root element.
        """
        elem = root.find('main:workbookProtection', namespaces=self.ns)
        if elem is None:
            return

        protection._lock_structure = elem.get('lockStructure') == '1'
        protection._lock_windows = elem.get('lockWindows') == '1'
        protection._lock_revision = elem.get('lockRevision') == '1'
        if elem.get('workbookPassword'):
            protection._workbook_password = elem.get('workbookPassword')
        if elem.get('revisionsPassword'):
            protection._revisions_password = elem.get('revisionsPassword')

    def load_workbook_pr(self, workbook_pr, root):
        """
        Loads workbook properties (workbookPr) from XML.

        ECMA-376 Section: 18.2.13

        Args:
            workbook_pr: WorkbookPr object to load into.
            root: XML root element.
        """
        elem = root.find('main:workbookPr', namespaces=self.ns)
        if elem is None:
            return

        workbook_pr._date1904 = elem.get('date1904') == '1'
        if elem.get('codeName'):
            workbook_pr._code_name = elem.get('codeName')
        if elem.get('showObjects'):
            workbook_pr._show_objects = elem.get('showObjects')
        workbook_pr._filter_privacy = elem.get('filterPrivacy') == '1'
        if elem.get('showBorderUnselectedTables') is not None:
            workbook_pr._show_border_unselected_tables = elem.get('showBorderUnselectedTables') != '0'
        if elem.get('showInkAnnotation') is not None:
            workbook_pr._show_ink_annotation = elem.get('showInkAnnotation') != '0'
        workbook_pr._backup_file = elem.get('backupFile') == '1'
        if elem.get('saveExternalLinkValues') is not None:
            workbook_pr._save_external_link_values = elem.get('saveExternalLinkValues') != '0'
        if elem.get('updateLinks'):
            workbook_pr._update_links = elem.get('updateLinks')
        workbook_pr._hide_pivot_field_list = elem.get('hidePivotFieldList') == '1'
        if elem.get('defaultThemeVersion'):
            workbook_pr._default_theme_version = int(elem.get('defaultThemeVersion'))

    def load_book_views(self, view, root):
        """
        Loads book views from XML.

        ECMA-376 Section: 18.2.1, 18.2.30

        Args:
            view: WorkbookView object to load into.
            root: XML root element.
        """
        book_views = root.find('main:bookViews', namespaces=self.ns)
        if book_views is None:
            return

        wb_view = book_views.find('main:workbookView', namespaces=self.ns)
        if wb_view is None:
            return

        if wb_view.get('xWindow'):
            view._x_window = int(wb_view.get('xWindow'))
        if wb_view.get('yWindow'):
            view._y_window = int(wb_view.get('yWindow'))
        if wb_view.get('windowWidth'):
            view._window_width = int(wb_view.get('windowWidth'))
        if wb_view.get('windowHeight'):
            view._window_height = int(wb_view.get('windowHeight'))
        if wb_view.get('activeTab'):
            view._active_tab = int(wb_view.get('activeTab'))
        if wb_view.get('firstSheet'):
            view._first_sheet = int(wb_view.get('firstSheet'))

        view._show_horizontal_scroll = wb_view.get('showHorizontalScroll', '1') != '0'
        view._show_vertical_scroll = wb_view.get('showVerticalScroll', '1') != '0'
        view._show_sheet_tabs = wb_view.get('showSheetTabs', '1') != '0'

        if wb_view.get('tabRatio'):
            view._tab_ratio = int(wb_view.get('tabRatio'))
        view._minimized = wb_view.get('minimized') == '1'
        if wb_view.get('visibility'):
            view._visibility = wb_view.get('visibility')

    def load_calc_pr(self, calculation, root):
        """
        Loads calculation properties from XML.

        ECMA-376 Section: 18.2.2

        Args:
            calculation: CalculationProperties object to load into.
            root: XML root element.
        """
        elem = root.find('main:calcPr', namespaces=self.ns)
        if elem is None:
            return

        if elem.get('calcId'):
            calculation._calc_id = int(elem.get('calcId'))
        if elem.get('calcMode'):
            calculation._calc_mode = elem.get('calcMode')
        calculation._full_calc_on_load = elem.get('fullCalcOnLoad') == '1'
        if elem.get('refMode'):
            calculation._ref_mode = elem.get('refMode')
        calculation._iterate = elem.get('iterate') == '1'
        if elem.get('iterateCount'):
            calculation._iterate_count = int(elem.get('iterateCount'))
        if elem.get('iterateDelta'):
            calculation._iterate_delta = float(elem.get('iterateDelta'))
        calculation._full_precision = elem.get('fullPrecision', '1') != '0'
        calculation._calc_on_save = elem.get('calcOnSave', '1') != '0'
        calculation._concurrent_calc = elem.get('concurrentCalc', '1') != '0'
        calculation._calc_completed = elem.get('calcCompleted', '1') != '0'
        calculation._force_full_calc = elem.get('forceFullCalc') == '1'

    def load_defined_names(self, defined_names, root):
        """
        Loads defined names from XML.

        ECMA-376 Section: 18.2.6

        Args:
            defined_names: DefinedNameCollection object to load into.
            root: XML root element.
        """
        from .workbook_properties import DefinedName

        dn_container = root.find('main:definedNames', namespaces=self.ns)
        if dn_container is None:
            return

        for dn_elem in dn_container.findall('main:definedName', namespaces=self.ns):
            name = dn_elem.get('name')
            refers_to = dn_elem.text or ''

            local_sheet_id = None
            if dn_elem.get('localSheetId'):
                local_sheet_id = int(dn_elem.get('localSheetId'))

            dn = DefinedName(name, refers_to, local_sheet_id)
            dn._hidden = dn_elem.get('hidden') == '1'
            dn._comment = dn_elem.get('comment')

            defined_names._names.append(dn)


class WorksheetPropertiesXMLLoader:
    """
    Handles loading worksheet properties from XML format for .xlsx files.

    Examples:
        >>> loader = WorksheetPropertiesXMLLoader(namespaces)
        >>> loader.load_sheet_views(worksheet.properties, root)
    """

    def __init__(self, namespaces):
        """
        Initializes a new instance of the WorksheetPropertiesXMLLoader class.

        Args:
            namespaces: XML namespaces dictionary for parsing.
        """
        self.ns = namespaces

    def load_sheet_views(self, properties, root):
        """
        Loads sheet views from XML.

        ECMA-376 Section: 18.3.1.88, 18.3.1.87

        Args:
            properties: WorksheetProperties object to load into.
            root: XML root element.
        """
        sheet_views = root.find('main:sheetViews', namespaces=self.ns)
        if sheet_views is None:
            return

        sheet_view = sheet_views.find('main:sheetView', namespaces=self.ns)
        if sheet_view is None:
            return

        view = properties.view
        view._show_formulas = sheet_view.get('showFormulas') == '1'
        view._show_grid_lines = sheet_view.get('showGridLines', '1') != '0'
        view._show_row_col_headers = sheet_view.get('showRowColHeaders', '1') != '0'
        view._show_zeros = sheet_view.get('showZeros', '1') != '0'
        view._right_to_left = sheet_view.get('rightToLeft') == '1'
        view._tab_selected = sheet_view.get('tabSelected') == '1'
        view._show_ruler = sheet_view.get('showRuler', '1') != '0'
        view._show_outline_symbols = sheet_view.get('showOutlineSymbols', '1') != '0'
        view._default_grid_color = sheet_view.get('defaultGridColor', '1') != '0'
        view._show_white_space = sheet_view.get('showWhiteSpace', '1') != '0'
        view._window_protection = sheet_view.get('windowProtection') == '1'

        if sheet_view.get('view'):
            view._view = sheet_view.get('view')
        if sheet_view.get('topLeftCell'):
            view._top_left_cell = sheet_view.get('topLeftCell')
        if sheet_view.get('colorId'):
            view._color_id = int(sheet_view.get('colorId'))
        if sheet_view.get('zoomScale'):
            view._zoom_scale = int(sheet_view.get('zoomScale'))
        if sheet_view.get('zoomScaleNormal'):
            view._zoom_scale_normal = int(sheet_view.get('zoomScaleNormal'))
        if sheet_view.get('zoomScaleSheetLayoutView'):
            view._zoom_scale_sheet_layout_view = int(sheet_view.get('zoomScaleSheetLayoutView'))
        if sheet_view.get('zoomScalePageLayoutView'):
            view._zoom_scale_page_layout_view = int(sheet_view.get('zoomScalePageLayoutView'))
        if sheet_view.get('workbookViewId'):
            view._workbook_view_id = int(sheet_view.get('workbookViewId'))

        # Load pane
        pane_elem = sheet_view.find('main:pane', namespaces=self.ns)
        if pane_elem is not None:
            pane = properties.pane
            if pane_elem.get('xSplit'):
                pane._x_split = float(pane_elem.get('xSplit'))
            if pane_elem.get('ySplit'):
                pane._y_split = float(pane_elem.get('ySplit'))
            if pane_elem.get('topLeftCell'):
                pane._top_left_cell = pane_elem.get('topLeftCell')
            if pane_elem.get('activePane'):
                pane._active_pane = pane_elem.get('activePane')
            if pane_elem.get('state'):
                pane._state = pane_elem.get('state')

        # Load selection
        selection_elem = sheet_view.find('main:selection', namespaces=self.ns)
        if selection_elem is not None:
            selection = properties.selection
            if selection_elem.get('pane'):
                selection._pane = selection_elem.get('pane')
            if selection_elem.get('activeCell'):
                selection._active_cell = selection_elem.get('activeCell')
            if selection_elem.get('sqref'):
                selection._sqref = selection_elem.get('sqref')

    def load_sheet_format_pr(self, properties, root):
        """
        Loads sheet format properties from XML.

        ECMA-376 Section: 18.3.1.82

        Args:
            properties: WorksheetProperties object to load into.
            root: XML root element.
        """
        elem = root.find('main:sheetFormatPr', namespaces=self.ns)
        if elem is None:
            return

        format_props = properties.format
        if elem.get('baseColWidth'):
            format_props._base_col_width = int(elem.get('baseColWidth'))
        if elem.get('defaultColWidth'):
            format_props._default_col_width = float(elem.get('defaultColWidth'))
        if elem.get('defaultRowHeight'):
            format_props._default_row_height = float(elem.get('defaultRowHeight'))
        format_props._custom_height = elem.get('customHeight') == '1'
        format_props._zero_height = elem.get('zeroHeight') == '1'
        format_props._thick_top = elem.get('thickTop') == '1'
        format_props._thick_bottom = elem.get('thickBottom') == '1'
        if elem.get('outlineLevelRow'):
            format_props._outline_level_row = int(elem.get('outlineLevelRow'))
        if elem.get('outlineLevelCol'):
            format_props._outline_level_col = int(elem.get('outlineLevelCol'))

    def load_sheet_protection(self, properties, root):
        """
        Loads sheet protection from XML.

        ECMA-376 Section: 18.3.1.85

        Args:
            properties: WorksheetProperties object to load into.
            root: XML root element.
        """
        elem = root.find('main:sheetProtection', namespaces=self.ns)
        if elem is None:
            return

        prot = properties.protection
        prot._sheet = elem.get('sheet') == '1'
        prot._objects = elem.get('objects') == '1'
        prot._scenarios = elem.get('scenarios') == '1'
        prot._format_cells = elem.get('formatCells', '1') != '0'
        prot._format_columns = elem.get('formatColumns', '1') != '0'
        prot._format_rows = elem.get('formatRows', '1') != '0'
        prot._insert_columns = elem.get('insertColumns', '1') != '0'
        prot._insert_rows = elem.get('insertRows', '1') != '0'
        prot._insert_hyperlinks = elem.get('insertHyperlinks', '1') != '0'
        prot._delete_columns = elem.get('deleteColumns', '1') != '0'
        prot._delete_rows = elem.get('deleteRows', '1') != '0'
        prot._select_locked_cells = elem.get('selectLockedCells') == '1'
        prot._sort = elem.get('sort', '1') != '0'
        prot._auto_filter = elem.get('autoFilter', '1') != '0'
        prot._pivot_tables = elem.get('pivotTables', '1') != '0'
        prot._select_unlocked_cells = elem.get('selectUnlockedCells') == '1'
        if elem.get('password'):
            prot._password = elem.get('password')
        if elem.get('algorithmName'):
            prot._algorithm_name = elem.get('algorithmName')
        if elem.get('hashValue'):
            prot._hash_value = elem.get('hashValue')
        if elem.get('saltValue'):
            prot._salt_value = elem.get('saltValue')
        if elem.get('spinCount'):
            prot._spin_count = int(elem.get('spinCount'))

    def load_print_options(self, properties, root):
        """
        Loads print options from XML.

        ECMA-376 Section: 18.3.1.70

        Args:
            properties: WorksheetProperties object to load into.
            root: XML root element.
        """
        elem = root.find('main:printOptions', namespaces=self.ns)
        if elem is None:
            return

        print_opts = properties.print_options
        print_opts._print_headings = elem.get('headings') == '1'
        print_opts._print_grid_lines = elem.get('gridLines') == '1'
        print_opts._horizontal_centered = elem.get('horizontalCentered') == '1'
        print_opts._vertical_centered = elem.get('verticalCentered') == '1'
        print_opts._black_and_white = elem.get('blackAndWhite') == '1'
        print_opts._draft_quality = elem.get('draft') == '1'
        if elem.get('comments'):
            print_opts._cell_comments = elem.get('comments')
        if elem.get('errors'):
            print_opts._print_errors = elem.get('errors')

    def load_page_margins(self, properties, root):
        """
        Loads page margins from XML.

        ECMA-376 Section: 18.3.1.62

        Args:
            properties: WorksheetProperties object to load into.
            root: XML root element.
        """
        elem = root.find('main:pageMargins', namespaces=self.ns)
        if elem is None:
            return

        margins = properties.page_margins
        if elem.get('left'):
            margins._left = float(elem.get('left'))
        if elem.get('right'):
            margins._right = float(elem.get('right'))
        if elem.get('top'):
            margins._top = float(elem.get('top'))
        if elem.get('bottom'):
            margins._bottom = float(elem.get('bottom'))
        if elem.get('header'):
            margins._header = float(elem.get('header'))
        if elem.get('footer'):
            margins._footer = float(elem.get('footer'))

    def load_page_setup(self, properties, root):
        """
        Loads page setup from XML.

        ECMA-376 Section: 18.3.1.63

        Args:
            properties: WorksheetProperties object to load into.
            root: XML root element.
        """
        elem = root.find('main:pageSetup', namespaces=self.ns)
        if elem is None:
            return

        page_setup = properties.page_setup
        if elem.get('paperSize'):
            page_setup._paper_size = int(elem.get('paperSize'))
        if elem.get('scale'):
            page_setup._scale = int(elem.get('scale'))
        if elem.get('firstPageNumber'):
            page_setup._first_page_number = int(elem.get('firstPageNumber'))
        if elem.get('fitToWidth'):
            page_setup._fit_to_width = int(elem.get('fitToWidth'))
        if elem.get('fitToHeight'):
            page_setup._fit_to_height = int(elem.get('fitToHeight'))
        if elem.get('pageOrder'):
            page_setup._page_order = elem.get('pageOrder')
        if elem.get('orientation'):
            page_setup._orientation = elem.get('orientation')
        if elem.get('usePrinterDefaults') is not None:
            page_setup._use_printer_defaults = elem.get('usePrinterDefaults') != '0'
        page_setup._black_and_white = elem.get('blackAndWhite') == '1'
        page_setup._draft = elem.get('draft') == '1'
        if elem.get('comments'):
            page_setup._cell_comments = elem.get('comments')
        if elem.get('useFirstPageNumber') == '1' and page_setup._first_page_number is None:
            page_setup._first_page_number = 1
        if elem.get('errors'):
            page_setup._errors = elem.get('errors')
        if elem.get('horizontalDpi'):
            page_setup._horizontal_dpi = int(elem.get('horizontalDpi'))
        if elem.get('verticalDpi'):
            page_setup._vertical_dpi = int(elem.get('verticalDpi'))
        if elem.get('copies'):
            page_setup._copies = int(elem.get('copies'))

    def load_header_footer(self, properties, root):
        """
        Loads header/footer from XML.

        ECMA-376 Section: 18.3.1.46

        Args:
            properties: WorksheetProperties object to load into.
            root: XML root element.
        """
        elem = root.find('main:headerFooter', namespaces=self.ns)
        if elem is None:
            return

        hf = properties.header_footer
        hf._different_first = elem.get('differentFirst') == '1'
        hf._different_odd_even = elem.get('differentOddEven') == '1'
        hf._scale_with_doc = elem.get('scaleWithDoc', '1') != '0'
        hf._align_with_margins = elem.get('alignWithMargins', '1') != '0'

        odd_header = elem.find('main:oddHeader', namespaces=self.ns)
        if odd_header is not None and odd_header.text:
            hf._odd_header = odd_header.text
        odd_footer = elem.find('main:oddFooter', namespaces=self.ns)
        if odd_footer is not None and odd_footer.text:
            hf._odd_footer = odd_footer.text
        even_header = elem.find('main:evenHeader', namespaces=self.ns)
        if even_header is not None and even_header.text:
            hf._even_header = even_header.text
        even_footer = elem.find('main:evenFooter', namespaces=self.ns)
        if even_footer is not None and even_footer.text:
            hf._even_footer = even_footer.text
        first_header = elem.find('main:firstHeader', namespaces=self.ns)
        if first_header is not None and first_header.text:
            hf._first_header = first_header.text
        first_footer = elem.find('main:firstFooter', namespaces=self.ns)
        if first_footer is not None and first_footer.text:
            hf._first_footer = first_footer.text
