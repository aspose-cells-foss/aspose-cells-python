"""
Aspose.Cells for Python - Worksheet Properties Module

This module provides classes for worksheet-level properties according to ECMA-376 specification.
Includes sheet views, sheet format, sheet protection, page setup, page margins, and header/footer.

ECMA-376 Sections: 18.3.1
"""


class SheetView:
    """
    Represents a sheet view configuration.

    ECMA-376 Section: 18.3.1.87

    Examples:
        >>> ws.properties.view.show_grid_lines = False
        >>> ws.properties.view.zoom_scale = 80
    """

    def __init__(self):
        self._window_protection = False
        self._show_formulas = False
        self._show_grid_lines = True
        self._show_row_col_headers = True
        self._show_outline_symbols = True
        self._show_zeros = True
        self._right_to_left = False
        self._tab_selected = False
        self._show_ruler = True
        self._show_white_space = True
        self._view = 'normal'
        self._top_left_cell = None
        self._default_grid_color = True
        self._color_id = None
        self._zoom_scale = 100
        self._zoom_scale_normal = 0
        self._zoom_scale_page_layout_view = 0
        self._zoom_scale_sheet_layout_view = 0
        self._workbook_view_id = 0

    @property
    def show_formulas(self):
        """Whether to show formulas instead of values."""
        return self._show_formulas

    @show_formulas.setter
    def show_formulas(self, value):
        self._show_formulas = value

    @property
    def show_grid_lines(self):
        """Whether to show grid lines."""
        return self._show_grid_lines

    @show_grid_lines.setter
    def show_grid_lines(self, value):
        self._show_grid_lines = value

    @property
    def show_row_col_headers(self):
        """Whether to show row and column headers."""
        return self._show_row_col_headers

    @show_row_col_headers.setter
    def show_row_col_headers(self, value):
        self._show_row_col_headers = value

    @property
    def show_zeros(self):
        """Whether to display zero values."""
        return self._show_zeros

    @show_zeros.setter
    def show_zeros(self, value):
        self._show_zeros = value

    @property
    def right_to_left(self):
        """Whether to display right-to-left."""
        return self._right_to_left

    @right_to_left.setter
    def right_to_left(self, value):
        self._right_to_left = value

    @property
    def tab_selected(self):
        """Whether this sheet is selected."""
        return self._tab_selected

    @tab_selected.setter
    def tab_selected(self, value):
        self._tab_selected = value

    @property
    def view(self):
        """View mode: 'normal', 'pageBreakPreview', or 'pageLayout'."""
        return self._view

    @view.setter
    def view(self, value):
        if value not in ('normal', 'pageBreakPreview', 'pageLayout'):
            raise ValueError("view must be 'normal', 'pageBreakPreview', or 'pageLayout'")
        self._view = value

    @property
    def top_left_cell(self):
        """Top-left visible cell."""
        return self._top_left_cell

    @top_left_cell.setter
    def top_left_cell(self, value):
        self._top_left_cell = value

    @property
    def zoom_scale(self):
        """Zoom scale (10-400, default: 100)."""
        return self._zoom_scale

    @zoom_scale.setter
    def zoom_scale(self, value):
        if value < 10 or value > 400:
            raise ValueError("zoom_scale must be between 10 and 400")
        self._zoom_scale = value

    @property
    def zoom_scale_normal(self):
        """Normal zoom scale."""
        return self._zoom_scale_normal

    @zoom_scale_normal.setter
    def zoom_scale_normal(self, value):
        self._zoom_scale_normal = value

    @property
    def zoom_scale_page_layout_view(self):
        """Zoom for page layout view."""
        return self._zoom_scale_page_layout_view

    @zoom_scale_page_layout_view.setter
    def zoom_scale_page_layout_view(self, value):
        self._zoom_scale_page_layout_view = value

    @property
    def view_type(self):
        """Alias for view property. View mode: 'normal', 'pageBreakPreview', or 'pageLayout'."""
        return self._view

    @view_type.setter
    def view_type(self, value):
        if value not in ('normal', 'pageBreakPreview', 'pageLayout'):
            raise ValueError("view_type must be 'normal', 'pageBreakPreview', or 'pageLayout'")
        self._view = value

    @property
    def window_protection(self):
        """Whether the sheet window is protected."""
        return self._window_protection

    @window_protection.setter
    def window_protection(self, value):
        self._window_protection = value

    @property
    def show_outline_symbols(self):
        """Whether to show outline symbols."""
        return self._show_outline_symbols

    @show_outline_symbols.setter
    def show_outline_symbols(self, value):
        self._show_outline_symbols = value

    @property
    def show_ruler(self):
        """Whether to show the ruler."""
        return self._show_ruler

    @show_ruler.setter
    def show_ruler(self, value):
        self._show_ruler = value

    @property
    def show_white_space(self):
        """Whether to show white space."""
        return self._show_white_space

    @show_white_space.setter
    def show_white_space(self, value):
        self._show_white_space = value

    @property
    def default_grid_color(self):
        """Whether to use the default grid color."""
        return self._default_grid_color

    @default_grid_color.setter
    def default_grid_color(self, value):
        self._default_grid_color = value

    @property
    def color_id(self):
        """Grid color index (default 64)."""
        return self._color_id

    @color_id.setter
    def color_id(self, value):
        self._color_id = value

    @property
    def zoom_scale_sheet_layout_view(self):
        """Zoom for sheet layout view."""
        return self._zoom_scale_sheet_layout_view

    @zoom_scale_sheet_layout_view.setter
    def zoom_scale_sheet_layout_view(self, value):
        self._zoom_scale_sheet_layout_view = value

    @property
    def workbook_view_id(self):
        """Workbook view index for this sheet view."""
        return self._workbook_view_id

    @workbook_view_id.setter
    def workbook_view_id(self, value):
        self._workbook_view_id = value


class Selection:
    """
    Represents cell selection in a sheet view.

    ECMA-376 Section: 18.3.1.81

    Examples:
        >>> ws.properties.selection.active_cell = "B2"
        >>> ws.properties.selection.sqref = "B2:D5"
    """

    def __init__(self):
        self._pane = None
        self._active_cell = 'A1'
        self._sqref = 'A1'
        self._active_cell_id = 0

    @property
    def pane(self):
        """Pane selection is in: 'topLeft', 'topRight', 'bottomLeft', 'bottomRight'."""
        return self._pane

    @pane.setter
    def pane(self, value):
        self._pane = value

    @property
    def active_cell(self):
        """Active cell reference."""
        return self._active_cell

    @active_cell.setter
    def active_cell(self, value):
        self._active_cell = value

    @property
    def sqref(self):
        """Selected range (space-delimited for multiple ranges)."""
        return self._sqref

    @sqref.setter
    def sqref(self, value):
        self._sqref = value


class Pane:
    """
    Represents pane (freeze/split) settings.

    ECMA-376 Section: 18.3.1.66

    Examples:
        >>> ws.properties.pane.x_split = 1
        >>> ws.properties.pane.y_split = 1
        >>> ws.properties.pane.state = "frozen"
    """

    def __init__(self):
        self._x_split = None
        self._y_split = None
        self._top_left_cell = None
        self._active_pane = None
        self._state = None

    @property
    def x_split(self):
        """Horizontal split position (column count or pixel position)."""
        return self._x_split

    @x_split.setter
    def x_split(self, value):
        self._x_split = value

    @property
    def y_split(self):
        """Vertical split position (row count or pixel position)."""
        return self._y_split

    @y_split.setter
    def y_split(self, value):
        self._y_split = value

    @property
    def top_left_cell(self):
        """Top-left cell of bottom-right pane."""
        return self._top_left_cell

    @top_left_cell.setter
    def top_left_cell(self, value):
        self._top_left_cell = value

    @property
    def active_pane(self):
        """Active pane: 'bottomLeft', 'bottomRight', 'topLeft', 'topRight'."""
        return self._active_pane

    @active_pane.setter
    def active_pane(self, value):
        self._active_pane = value

    @property
    def state(self):
        """Pane state: 'frozen', 'frozenSplit', or 'split'."""
        return self._state

    @state.setter
    def state(self, value):
        if value is not None and value not in ('frozen', 'frozenSplit', 'split'):
            raise ValueError("state must be 'frozen', 'frozenSplit', or 'split'")
        self._state = value

    @property
    def is_frozen(self):
        """Returns True if pane is frozen."""
        return self._state in ('frozen', 'frozenSplit')


class SheetFormatProperties:
    """
    Represents sheet format properties.

    ECMA-376 Section: 18.3.1.82

    Examples:
        >>> ws.properties.format.default_row_height = 15
        >>> ws.properties.format.default_col_width = 8.43
    """

    def __init__(self):
        self._base_col_width = 8
        self._default_col_width = None
        self._default_row_height = 15.0
        self._custom_height = False
        self._zero_height = False
        self._thick_top = False
        self._thick_bottom = False
        self._outline_level_row = 0
        self._outline_level_col = 0

    @property
    def base_col_width(self):
        """Base column width in characters (default: 8)."""
        return self._base_col_width

    @base_col_width.setter
    def base_col_width(self, value):
        self._base_col_width = value

    @property
    def default_col_width(self):
        """Default column width in characters."""
        return self._default_col_width

    @default_col_width.setter
    def default_col_width(self, value):
        self._default_col_width = value

    @property
    def default_row_height(self):
        """Default row height in points."""
        return self._default_row_height

    @default_row_height.setter
    def default_row_height(self, value):
        self._default_row_height = value

    @property
    def custom_height(self):
        """Whether custom row height is applied."""
        return self._custom_height

    @custom_height.setter
    def custom_height(self, value):
        self._custom_height = value

    @property
    def zero_height(self):
        """Whether rows have zero height by default."""
        return self._zero_height

    @zero_height.setter
    def zero_height(self, value):
        self._zero_height = value

    @property
    def outline_level_row(self):
        """Default outline level for rows."""
        return self._outline_level_row

    @outline_level_row.setter
    def outline_level_row(self, value):
        self._outline_level_row = value

    @property
    def outline_level_col(self):
        """Default outline level for columns."""
        return self._outline_level_col

    @outline_level_col.setter
    def outline_level_col(self, value):
        self._outline_level_col = value

    @property
    def thick_top(self):
        """Whether thick top border is applied."""
        return self._thick_top

    @thick_top.setter
    def thick_top(self, value):
        self._thick_top = value

    @property
    def thick_bottom(self):
        """Whether thick bottom border is applied."""
        return self._thick_bottom

    @thick_bottom.setter
    def thick_bottom(self, value):
        self._thick_bottom = value


class SheetProtection:
    """
    Represents sheet protection settings.

    ECMA-376 Section: 18.3.1.85

    Examples:
        >>> ws.properties.protection.sheet = True
        >>> ws.properties.protection.format_cells = False
    """

    def __init__(self):
        self._algorithm_name = None
        self._hash_value = None
        self._salt_value = None
        self._spin_count = None
        self._sheet = False
        self._objects = False
        self._scenarios = False
        self._format_cells = True
        self._format_columns = True
        self._format_rows = True
        self._insert_columns = True
        self._insert_rows = True
        self._insert_hyperlinks = True
        self._delete_columns = True
        self._delete_rows = True
        self._select_locked_cells = False
        self._sort = True
        self._auto_filter = True
        self._pivot_tables = True
        self._select_unlocked_cells = False
        self._password = None

    @property
    def sheet(self):
        """Whether to protect sheet."""
        return self._sheet

    @sheet.setter
    def sheet(self, value):
        self._sheet = value

    @property
    def objects(self):
        """Whether to protect objects."""
        return self._objects

    @objects.setter
    def objects(self, value):
        self._objects = value

    @property
    def scenarios(self):
        """Whether to protect scenarios."""
        return self._scenarios

    @scenarios.setter
    def scenarios(self, value):
        self._scenarios = value

    @property
    def format_cells(self):
        """Whether to allow formatting cells."""
        return self._format_cells

    @format_cells.setter
    def format_cells(self, value):
        self._format_cells = value

    @property
    def format_columns(self):
        """Whether to allow formatting columns."""
        return self._format_columns

    @format_columns.setter
    def format_columns(self, value):
        self._format_columns = value

    @property
    def format_rows(self):
        """Whether to allow formatting rows."""
        return self._format_rows

    @format_rows.setter
    def format_rows(self, value):
        self._format_rows = value

    @property
    def insert_columns(self):
        """Whether to allow inserting columns."""
        return self._insert_columns

    @insert_columns.setter
    def insert_columns(self, value):
        self._insert_columns = value

    @property
    def insert_rows(self):
        """Whether to allow inserting rows."""
        return self._insert_rows

    @insert_rows.setter
    def insert_rows(self, value):
        self._insert_rows = value

    @property
    def insert_hyperlinks(self):
        """Whether to allow inserting hyperlinks."""
        return self._insert_hyperlinks

    @insert_hyperlinks.setter
    def insert_hyperlinks(self, value):
        self._insert_hyperlinks = value

    @property
    def delete_columns(self):
        """Whether to allow deleting columns."""
        return self._delete_columns

    @delete_columns.setter
    def delete_columns(self, value):
        self._delete_columns = value

    @property
    def delete_rows(self):
        """Whether to allow deleting rows."""
        return self._delete_rows

    @delete_rows.setter
    def delete_rows(self, value):
        self._delete_rows = value

    @property
    def select_locked_cells(self):
        """Whether to allow selecting locked cells."""
        return self._select_locked_cells

    @select_locked_cells.setter
    def select_locked_cells(self, value):
        self._select_locked_cells = value

    @property
    def sort(self):
        """Whether to allow sorting."""
        return self._sort

    @sort.setter
    def sort(self, value):
        self._sort = value

    @property
    def auto_filter(self):
        """Whether to allow using autoFilter."""
        return self._auto_filter

    @auto_filter.setter
    def auto_filter(self, value):
        self._auto_filter = value

    @property
    def pivot_tables(self):
        """Whether to allow using pivot tables."""
        return self._pivot_tables

    @pivot_tables.setter
    def pivot_tables(self, value):
        self._pivot_tables = value

    @property
    def select_unlocked_cells(self):
        """Whether to allow selecting unlocked cells."""
        return self._select_unlocked_cells

    @select_unlocked_cells.setter
    def select_unlocked_cells(self, value):
        self._select_unlocked_cells = value

    @property
    def password(self):
        """Protection password (hashed)."""
        return self._password

    @password.setter
    def password(self, value):
        self._password = value

    @property
    def algorithm_name(self):
        """Hash algorithm name (for modern protection)."""
        return self._algorithm_name

    @algorithm_name.setter
    def algorithm_name(self, value):
        self._algorithm_name = value

    @property
    def hash_value(self):
        """Password hash (base64)."""
        return self._hash_value

    @hash_value.setter
    def hash_value(self, value):
        self._hash_value = value

    @property
    def salt_value(self):
        """Salt for hashing (base64)."""
        return self._salt_value

    @salt_value.setter
    def salt_value(self, value):
        self._salt_value = value

    @property
    def spin_count(self):
        """Number of hash iterations."""
        return self._spin_count

    @spin_count.setter
    def spin_count(self, value):
        self._spin_count = value

    @property
    def is_protected(self):
        """Returns True if sheet protection is enabled."""
        return self._sheet


class PageSetup:
    """
    Represents page setup settings.

    ECMA-376 Section: 18.3.1.63

    Examples:
        >>> ws.properties.page_setup.orientation = "landscape"
        >>> ws.properties.page_setup.paper_size = 9  # A4
        >>> ws.properties.page_setup.scale = 80
    """

    def __init__(self):
        self._paper_size = 1  # Letter
        self._scale = 100
        self._first_page_number = None
        self._fit_to_width = None
        self._fit_to_height = None
        self._page_order = 'downThenOver'
        self._orientation = 'portrait'
        self._use_printer_defaults = True
        self._black_and_white = False
        self._draft = False
        self._cell_comments = 'none'
        self._errors = 'displayed'
        self._horizontal_dpi = None
        self._vertical_dpi = None
        self._copies = 1

    @property
    def paper_size(self):
        """Paper size (1=Letter, 9=A4, etc.)."""
        return self._paper_size

    @paper_size.setter
    def paper_size(self, value):
        self._paper_size = value

    @property
    def scale(self):
        """Print scaling (10-400)."""
        return self._scale

    @scale.setter
    def scale(self, value):
        if value < 10 or value > 400:
            raise ValueError("scale must be between 10 and 400")
        self._scale = value

    @property
    def first_page_number(self):
        """First page number."""
        return self._first_page_number

    @first_page_number.setter
    def first_page_number(self, value):
        self._first_page_number = value

    @property
    def fit_to_width(self):
        """Fit to width pages (0 = use scale)."""
        return self._fit_to_width

    @fit_to_width.setter
    def fit_to_width(self, value):
        self._fit_to_width = value

    @property
    def fit_to_height(self):
        """Fit to height pages (0 = use scale)."""
        return self._fit_to_height

    @fit_to_height.setter
    def fit_to_height(self, value):
        self._fit_to_height = value

    @property
    def page_order(self):
        """Page order: 'downThenOver' or 'overThenDown'."""
        return self._page_order

    @page_order.setter
    def page_order(self, value):
        if value not in ('downThenOver', 'overThenDown'):
            raise ValueError("page_order must be 'downThenOver' or 'overThenDown'")
        self._page_order = value

    @property
    def orientation(self):
        """Page orientation: 'portrait' or 'landscape'."""
        return self._orientation

    @orientation.setter
    def orientation(self, value):
        if value not in ('portrait', 'landscape'):
            raise ValueError("orientation must be 'portrait' or 'landscape'")
        self._orientation = value

    @property
    def use_printer_defaults(self):
        """Whether to use printer defaults."""
        return self._use_printer_defaults

    @use_printer_defaults.setter
    def use_printer_defaults(self, value):
        self._use_printer_defaults = value

    @property
    def black_and_white(self):
        """Whether to print in black and white."""
        return self._black_and_white

    @black_and_white.setter
    def black_and_white(self, value):
        self._black_and_white = value

    @property
    def draft(self):
        """Whether to print in draft quality."""
        return self._draft

    @draft.setter
    def draft(self, value):
        self._draft = value

    @property
    def copies(self):
        """Number of copies."""
        return self._copies

    @copies.setter
    def copies(self, value):
        self._copies = value

    @property
    def cell_comments(self):
        """Print cell comments: 'none', 'atEnd', or 'asDisplayed'."""
        return self._cell_comments

    @cell_comments.setter
    def cell_comments(self, value):
        self._cell_comments = value

    @property
    def errors(self):
        """Print errors: 'displayed', 'blank', 'dash', or 'NA'."""
        return self._errors

    @errors.setter
    def errors(self, value):
        self._errors = value

    @property
    def horizontal_dpi(self):
        """Horizontal DPI for printing."""
        return self._horizontal_dpi

    @horizontal_dpi.setter
    def horizontal_dpi(self, value):
        self._horizontal_dpi = value

    @property
    def vertical_dpi(self):
        """Vertical DPI for printing."""
        return self._vertical_dpi

    @vertical_dpi.setter
    def vertical_dpi(self, value):
        self._vertical_dpi = value

    @property
    def use_first_page_number(self):
        """Whether to use the first page number."""
        return self._first_page_number is not None

    @use_first_page_number.setter
    def use_first_page_number(self, value):
        if value:
            if self._first_page_number is None:
                self._first_page_number = 1
        else:
            self._first_page_number = None

    @property
    def fit_to_page(self):
        """Whether to fit to page (True if fit_to_width or fit_to_height is set)."""
        return self._fit_to_width is not None or self._fit_to_height is not None

    @fit_to_page.setter
    def fit_to_page(self, value):
        """Enable or disable fit to page mode."""
        if value:
            # Enable fit to page with default values if not already set
            if self._fit_to_width is None:
                self._fit_to_width = 1
            if self._fit_to_height is None:
                self._fit_to_height = 1
        else:
            # Disable fit to page
            self._fit_to_width = None
            self._fit_to_height = None


class PageMargins:
    """
    Represents page margins.

    ECMA-376 Section: 18.3.1.62

    Examples:
        >>> ws.properties.page_margins.left = 0.7
        >>> ws.properties.page_margins.top = 0.75
    """

    def __init__(self):
        self._left = 0.7
        self._right = 0.7
        self._top = 0.75
        self._bottom = 0.75
        self._header = 0.3
        self._footer = 0.3

    @property
    def left(self):
        """Left margin in inches."""
        return self._left

    @left.setter
    def left(self, value):
        self._left = value

    @property
    def right(self):
        """Right margin in inches."""
        return self._right

    @right.setter
    def right(self, value):
        self._right = value

    @property
    def top(self):
        """Top margin in inches."""
        return self._top

    @top.setter
    def top(self, value):
        self._top = value

    @property
    def bottom(self):
        """Bottom margin in inches."""
        return self._bottom

    @bottom.setter
    def bottom(self, value):
        self._bottom = value

    @property
    def header(self):
        """Header margin in inches."""
        return self._header

    @header.setter
    def header(self, value):
        self._header = value

    @property
    def footer(self):
        """Footer margin in inches."""
        return self._footer

    @footer.setter
    def footer(self, value):
        self._footer = value


class HeaderFooter:
    """
    Represents header and footer settings.

    ECMA-376 Section: 18.3.1.46

    Examples:
        >>> ws.properties.header_footer.odd_header = "&C&\"Arial,Bold\"Report Title"
        >>> ws.properties.header_footer.odd_footer = "&CPage &P of &N"
    """

    def __init__(self):
        self._different_first = False
        self._different_odd_even = False
        self._scale_with_doc = True
        self._align_with_margins = True
        self._odd_header = None
        self._odd_footer = None
        self._even_header = None
        self._even_footer = None
        self._first_header = None
        self._first_footer = None

    @property
    def different_first(self):
        """Whether to have different first page header/footer."""
        return self._different_first

    @different_first.setter
    def different_first(self, value):
        self._different_first = value

    @property
    def different_odd_even(self):
        """Whether to have different odd/even page headers/footers."""
        return self._different_odd_even

    @different_odd_even.setter
    def different_odd_even(self, value):
        self._different_odd_even = value

    @property
    def scale_with_doc(self):
        """Whether to scale with document."""
        return self._scale_with_doc

    @scale_with_doc.setter
    def scale_with_doc(self, value):
        self._scale_with_doc = value

    @property
    def align_with_margins(self):
        """Whether to align with margins."""
        return self._align_with_margins

    @align_with_margins.setter
    def align_with_margins(self, value):
        self._align_with_margins = value

    @property
    def odd_header(self):
        """Header for odd pages (or all pages if not different_odd_even)."""
        return self._odd_header

    @odd_header.setter
    def odd_header(self, value):
        self._odd_header = value

    @property
    def odd_footer(self):
        """Footer for odd pages (or all pages if not different_odd_even)."""
        return self._odd_footer

    @odd_footer.setter
    def odd_footer(self, value):
        self._odd_footer = value

    @property
    def even_header(self):
        """Header for even pages."""
        return self._even_header

    @even_header.setter
    def even_header(self, value):
        self._even_header = value

    @property
    def even_footer(self):
        """Footer for even pages."""
        return self._even_footer

    @even_footer.setter
    def even_footer(self, value):
        self._even_footer = value

    @property
    def first_header(self):
        """Header for first page."""
        return self._first_header

    @first_header.setter
    def first_header(self, value):
        self._first_header = value

    @property
    def first_footer(self):
        """Footer for first page."""
        return self._first_footer

    @first_footer.setter
    def first_footer(self, value):
        self._first_footer = value


class PrintOptions:
    """
    Represents print options.

    ECMA-376 Section: 18.3.1.70

    Examples:
        >>> ws.properties.print_options.print_grid_lines = True
        >>> ws.properties.print_options.print_headings = True
    """

    def __init__(self):
        self._print_headings = False
        self._print_grid_lines = False
        self._grid_lines_set = True
        self._horizontal_centered = False
        self._vertical_centered = False
        self._black_and_white = False
        self._draft_quality = False
        self._cell_comments = 'none'
        self._print_errors = 'displayed'

    @property
    def print_headings(self):
        """Whether to print row and column headings."""
        return self._print_headings

    @print_headings.setter
    def print_headings(self, value):
        self._print_headings = value

    @property
    def print_grid_lines(self):
        """Whether to print grid lines."""
        return self._print_grid_lines

    @print_grid_lines.setter
    def print_grid_lines(self, value):
        self._print_grid_lines = value

    @property
    def horizontal_centered(self):
        """Whether to center horizontally on page."""
        return self._horizontal_centered

    @horizontal_centered.setter
    def horizontal_centered(self, value):
        self._horizontal_centered = value

    @property
    def vertical_centered(self):
        """Whether to center vertically on page."""
        return self._vertical_centered

    @vertical_centered.setter
    def vertical_centered(self, value):
        self._vertical_centered = value

    @property
    def print_gridlines(self):
        """Alias for print_grid_lines property."""
        return self._print_grid_lines

    @print_gridlines.setter
    def print_gridlines(self, value):
        self._print_grid_lines = value

    @property
    def black_and_white(self):
        """Whether to print in black and white."""
        return self._black_and_white

    @black_and_white.setter
    def black_and_white(self, value):
        self._black_and_white = value

    @property
    def draft_quality(self):
        """Whether to print in draft quality."""
        return self._draft_quality

    @draft_quality.setter
    def draft_quality(self, value):
        self._draft_quality = value

    @property
    def cell_comments(self):
        """Print cell comments: 'none', 'atEnd', or 'asDisplayed'."""
        return self._cell_comments

    @cell_comments.setter
    def cell_comments(self, value):
        self._cell_comments = value

    @property
    def print_errors(self):
        """Print errors: 'displayed', 'blank', 'dash', or 'NA'."""
        return self._print_errors

    @print_errors.setter
    def print_errors(self, value):
        self._print_errors = value


class WorksheetProperties:
    """
    Container for all worksheet-level properties.

    Examples:
        >>> ws.properties.view.show_grid_lines = False
        >>> ws.properties.protection.sheet = True
        >>> ws.properties.page_setup.orientation = "landscape"
    """

    def __init__(self):
        self._view = SheetView()
        self._selection = Selection()
        self._pane = Pane()
        self._format = SheetFormatProperties()
        self._protection = SheetProtection()
        self._page_setup = PageSetup()
        self._page_margins = PageMargins()
        self._header_footer = HeaderFooter()
        self._print_options = PrintOptions()

    @property
    def view(self):
        """Gets sheet view settings."""
        return self._view

    @property
    def selection(self):
        """Gets selection settings."""
        return self._selection

    @property
    def pane(self):
        """Gets pane (freeze/split) settings."""
        return self._pane

    @property
    def format(self):
        """Gets sheet format properties."""
        return self._format

    @property
    def protection(self):
        """Gets sheet protection settings."""
        return self._protection

    @property
    def page_setup(self):
        """Gets page setup settings."""
        return self._page_setup

    @property
    def page_margins(self):
        """Gets page margins."""
        return self._page_margins

    @property
    def header_footer(self):
        """Gets header and footer settings."""
        return self._header_footer

    @property
    def print_options(self):
        """Gets print options."""
        return self._print_options
