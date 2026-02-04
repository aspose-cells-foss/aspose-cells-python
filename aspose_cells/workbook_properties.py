"""
Aspose.Cells for Python - Workbook Properties Module

This module provides classes for workbook-level properties according to ECMA-376 specification.
Includes file version, workbook protection, workbook views, calculation properties, and defined names.

ECMA-376 Sections: 18.2.1-18.2.30
"""


class FileVersion:
    """
    Represents file version information for the workbook.

    ECMA-376 Section: 18.2.10

    Examples:
        >>> wb.properties.file_version.app_name = "xl"
        >>> wb.properties.file_version.last_edited = "7"
    """

    def __init__(self):
        self._app_name = "xl"
        self._last_edited = "7"
        self._lowest_edited = "0"
        self._rup_build = "12345"

    @property
    def app_name(self):
        """Application that created the file."""
        return self._app_name

    @app_name.setter
    def app_name(self, value):
        self._app_name = value

    @property
    def last_edited(self):
        """Number of times the file has been edited."""
        return self._last_edited

    @last_edited.setter
    def last_edited(self, value):
        self._last_edited = value

    @property
    def lowest_edited(self):
        """Lowest version of Excel that has edited the file."""
        return self._lowest_edited

    @lowest_edited.setter
    def lowest_edited(self, value):
        self._lowest_edited = value

    @property
    def rup_build(self):
        """Build number of the application."""
        return self._rup_build

    @rup_build.setter
    def rup_build(self, value):
        self._rup_build = value


class WorkbookPr:
    """
    Represents workbook properties (workbookPr element).

    ECMA-376 Section: 18.2.13

    Examples:
        >>> wb.properties.workbook_pr.date1904 = True
        >>> wb.properties.workbook_pr.code_name = "ThisWorkbook"
    """

    def __init__(self):
        self._date1904 = False
        self._code_name = None
        self._show_objects = "all"
        self._filter_privacy = False
        self._show_border_unselected_tables = True
        self._show_ink_annotation = True
        self._backup_file = False
        self._save_external_link_values = True
        self._update_links = "userSet"
        self._hide_pivot_field_list = False
        self._default_theme_version = None

    @property
    def date1904(self):
        """Whether to use 1904 date system."""
        return self._date1904

    @date1904.setter
    def date1904(self, value):
        self._date1904 = value

    @property
    def code_name(self):
        """VBA code name for the workbook."""
        return self._code_name

    @code_name.setter
    def code_name(self, value):
        self._code_name = value

    @property
    def show_objects(self):
        """Display objects mode: 'all', 'placeholders', or 'none'."""
        return self._show_objects

    @show_objects.setter
    def show_objects(self, value):
        self._show_objects = value

    @property
    def filter_privacy(self):
        """Whether to remove personal info from filters on save."""
        return self._filter_privacy

    @filter_privacy.setter
    def filter_privacy(self, value):
        self._filter_privacy = value

    @property
    def show_border_unselected_tables(self):
        """Whether to show borders for unselected tables."""
        return self._show_border_unselected_tables

    @show_border_unselected_tables.setter
    def show_border_unselected_tables(self, value):
        self._show_border_unselected_tables = value

    @property
    def show_ink_annotation(self):
        """Whether to show ink annotations."""
        return self._show_ink_annotation

    @show_ink_annotation.setter
    def show_ink_annotation(self, value):
        self._show_ink_annotation = value

    @property
    def backup_file(self):
        """Whether to create a backup file when saving."""
        return self._backup_file

    @backup_file.setter
    def backup_file(self, value):
        self._backup_file = value

    @property
    def save_external_link_values(self):
        """Whether to save cached values for external links."""
        return self._save_external_link_values

    @save_external_link_values.setter
    def save_external_link_values(self, value):
        self._save_external_link_values = value

    @property
    def update_links(self):
        """How to update external links: 'userSet', 'never', or 'always'."""
        return self._update_links

    @update_links.setter
    def update_links(self, value):
        self._update_links = value

    @property
    def hide_pivot_field_list(self):
        """Whether to hide the pivot field list."""
        return self._hide_pivot_field_list

    @hide_pivot_field_list.setter
    def hide_pivot_field_list(self, value):
        self._hide_pivot_field_list = value

    @property
    def default_theme_version(self):
        """Default theme version."""
        return self._default_theme_version

    @default_theme_version.setter
    def default_theme_version(self, value):
        self._default_theme_version = value


class WorkbookProtection:
    """
    Represents workbook protection settings.

    ECMA-376 Section: 18.2.29

    Examples:
        >>> wb.properties.protection.lock_structure = True
        >>> wb.properties.protection.workbook_password = "hashed_password"
    """

    def __init__(self):
        self._lock_structure = False
        self._lock_windows = False
        self._lock_revision = False
        self._revisions_password = None
        self._workbook_password = None

    @property
    def lock_structure(self):
        """Whether workbook structure is locked."""
        return self._lock_structure

    @lock_structure.setter
    def lock_structure(self, value):
        self._lock_structure = value

    @property
    def lock_windows(self):
        """Whether workbook windows are locked."""
        return self._lock_windows

    @lock_windows.setter
    def lock_windows(self, value):
        self._lock_windows = value

    @property
    def lock_revision(self):
        """Whether revision tracking is locked."""
        return self._lock_revision

    @lock_revision.setter
    def lock_revision(self, value):
        self._lock_revision = value

    @property
    def revisions_password(self):
        """Hashed password for revision protection."""
        return self._revisions_password

    @revisions_password.setter
    def revisions_password(self, value):
        self._revisions_password = value

    @property
    def workbook_password(self):
        """Hashed password for workbook protection."""
        return self._workbook_password

    @workbook_password.setter
    def workbook_password(self, value):
        self._workbook_password = value

    @property
    def is_protected(self):
        """Returns True if workbook has any protection enabled."""
        return self._lock_structure or self._lock_windows or self._lock_revision


class WorkbookView:
    """
    Represents a workbook view configuration.

    ECMA-376 Section: 18.2.30

    Examples:
        >>> wb.properties.view.active_tab = 0
        >>> wb.properties.view.show_sheet_tabs = True
    """

    def __init__(self):
        self._x_window = 0
        self._y_window = 0
        self._window_width = 22260
        self._window_height = 12645
        self._active_tab = 0
        self._first_sheet = 0
        self._show_horizontal_scroll = True
        self._show_vertical_scroll = True
        self._show_sheet_tabs = True
        self._tab_ratio = 600
        self._visibility = 'visible'
        self._minimized = False
        self._auto_filter_date_grouping = True

    @property
    def x_window(self):
        """Horizontal position of window."""
        return self._x_window

    @x_window.setter
    def x_window(self, value):
        self._x_window = value

    @property
    def y_window(self):
        """Vertical position of window."""
        return self._y_window

    @y_window.setter
    def y_window(self, value):
        self._y_window = value

    @property
    def window_width(self):
        """Width of workbook window."""
        return self._window_width

    @window_width.setter
    def window_width(self, value):
        self._window_width = value

    @property
    def window_height(self):
        """Height of workbook window."""
        return self._window_height

    @window_height.setter
    def window_height(self, value):
        self._window_height = value

    @property
    def active_tab(self):
        """Index of active sheet (0-based)."""
        return self._active_tab

    @active_tab.setter
    def active_tab(self, value):
        self._active_tab = value

    @property
    def first_sheet(self):
        """First sheet in the tab bar."""
        return self._first_sheet

    @first_sheet.setter
    def first_sheet(self, value):
        self._first_sheet = value

    @property
    def show_horizontal_scroll(self):
        """Whether to show horizontal scroll bar."""
        return self._show_horizontal_scroll

    @show_horizontal_scroll.setter
    def show_horizontal_scroll(self, value):
        self._show_horizontal_scroll = value

    @property
    def show_vertical_scroll(self):
        """Whether to show vertical scroll bar."""
        return self._show_vertical_scroll

    @show_vertical_scroll.setter
    def show_vertical_scroll(self, value):
        self._show_vertical_scroll = value

    @property
    def show_sheet_tabs(self):
        """Whether to show sheet tabs."""
        return self._show_sheet_tabs

    @show_sheet_tabs.setter
    def show_sheet_tabs(self, value):
        self._show_sheet_tabs = value

    @property
    def tab_ratio(self):
        """Ratio of tab bar width to horizontal split bar."""
        return self._tab_ratio

    @tab_ratio.setter
    def tab_ratio(self, value):
        self._tab_ratio = value

    @property
    def visibility(self):
        """Window visibility: 'visible', 'hidden', or 'veryHidden'."""
        return self._visibility

    @visibility.setter
    def visibility(self, value):
        self._visibility = value

    @property
    def minimized(self):
        """Whether window is minimized."""
        return self._minimized

    @minimized.setter
    def minimized(self, value):
        self._minimized = value


class CalculationProperties:
    """
    Represents calculation properties for the workbook.

    ECMA-376 Section: 18.2.2

    Examples:
        >>> wb.properties.calculation.calc_mode = "auto"
        >>> wb.properties.calculation.iterate = True
        >>> wb.properties.calculation.iterate_count = 100
    """

    def __init__(self):
        self._calc_id = None
        self._calc_mode = 'auto'
        self._full_calc_on_load = False
        self._ref_mode = 'A1'
        self._iterate = False
        self._iterate_count = 100
        self._iterate_delta = 0.001
        self._full_precision = True
        self._calc_completed = True
        self._calc_on_save = True
        self._concurrent_calc = True
        self._force_full_calc = False

    @property
    def calc_id(self):
        """Calculation engine version."""
        return self._calc_id

    @calc_id.setter
    def calc_id(self, value):
        self._calc_id = value

    @property
    def calc_mode(self):
        """Calculation mode: 'auto', 'manual', or 'autoNoTable'."""
        return self._calc_mode

    @calc_mode.setter
    def calc_mode(self, value):
        if value not in ('auto', 'manual', 'autoNoTable'):
            raise ValueError("calc_mode must be 'auto', 'manual', or 'autoNoTable'")
        self._calc_mode = value

    @property
    def full_calc_on_load(self):
        """Whether to recalculate on load."""
        return self._full_calc_on_load

    @full_calc_on_load.setter
    def full_calc_on_load(self, value):
        self._full_calc_on_load = value

    @property
    def ref_mode(self):
        """Reference style: 'A1' or 'R1C1'."""
        return self._ref_mode

    @ref_mode.setter
    def ref_mode(self, value):
        if value not in ('A1', 'R1C1'):
            raise ValueError("ref_mode must be 'A1' or 'R1C1'")
        self._ref_mode = value

    @property
    def iterate(self):
        """Whether to enable iterative calculation."""
        return self._iterate

    @iterate.setter
    def iterate(self, value):
        self._iterate = value

    @property
    def iterate_count(self):
        """Maximum iterations (default 100)."""
        return self._iterate_count

    @iterate_count.setter
    def iterate_count(self, value):
        self._iterate_count = value

    @property
    def iterate_delta(self):
        """Maximum change between iterations (default 0.001)."""
        return self._iterate_delta

    @iterate_delta.setter
    def iterate_delta(self, value):
        self._iterate_delta = value

    @property
    def full_precision(self):
        """Whether to use full precision."""
        return self._full_precision

    @full_precision.setter
    def full_precision(self, value):
        self._full_precision = value

    @property
    def calc_on_save(self):
        """Whether to recalculate before save."""
        return self._calc_on_save

    @calc_on_save.setter
    def calc_on_save(self, value):
        self._calc_on_save = value

    @property
    def concurrent_calc(self):
        """Whether to enable concurrent calculation."""
        return self._concurrent_calc

    @concurrent_calc.setter
    def concurrent_calc(self, value):
        self._concurrent_calc = value

    @property
    def calc_completed(self):
        """Whether the calculation is completed."""
        return self._calc_completed

    @calc_completed.setter
    def calc_completed(self, value):
        self._calc_completed = value

    @property
    def force_full_calc(self):
        """Whether to force a full calculation."""
        return self._force_full_calc

    @force_full_calc.setter
    def force_full_calc(self, value):
        self._force_full_calc = value


class DefinedName:
    """
    Represents a defined name in the workbook.

    ECMA-376 Section: 18.2.5

    Examples:
        >>> name = DefinedName("MyRange", "Sheet1!$A$1:$D$10")
        >>> wb.properties.defined_names.add(name)
    """

    def __init__(self, name, refers_to, local_sheet_id=None):
        self._name = name
        self._refers_to = refers_to
        self._local_sheet_id = local_sheet_id
        self._comment = None
        self._description = None
        self._hidden = False
        self._function = False
        self._vb_procedure = False

    @property
    def name(self):
        """Name of the defined name."""
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @property
    def refers_to(self):
        """Formula or range that the name refers to."""
        return self._refers_to

    @refers_to.setter
    def refers_to(self, value):
        self._refers_to = value

    @property
    def local_sheet_id(self):
        """Sheet index for sheet-local names (None for global names)."""
        return self._local_sheet_id

    @local_sheet_id.setter
    def local_sheet_id(self, value):
        self._local_sheet_id = value

    @property
    def comment(self):
        """Comment associated with the name."""
        return self._comment

    @comment.setter
    def comment(self, value):
        self._comment = value

    @property
    def description(self):
        """Description of the name."""
        return self._description

    @description.setter
    def description(self, value):
        self._description = value

    @property
    def hidden(self):
        """Whether the name is hidden."""
        return self._hidden

    @hidden.setter
    def hidden(self, value):
        self._hidden = value


class DefinedNameCollection:
    """
    Collection of defined names in the workbook.

    Examples:
        >>> wb.properties.defined_names.add(DefinedName("MyRange", "Sheet1!$A$1:$D$10"))
        >>> name = wb.properties.defined_names["MyRange"]
    """

    def __init__(self):
        self._names = []

    def add(self, name_or_str, refers_to=None, local_sheet_id=None):
        """
        Adds a defined name to the collection.

        Args:
            name_or_str: Either a DefinedName object or a string name.
            refers_to: Formula or range (required if name_or_str is a string).
            local_sheet_id: Sheet index for sheet-local names.

        Returns:
            DefinedName: The added defined name.
        """
        if isinstance(name_or_str, DefinedName):
            self._names.append(name_or_str)
            return name_or_str
        else:
            defined_name = DefinedName(name_or_str, refers_to, local_sheet_id)
            self._names.append(defined_name)
            return defined_name

    def remove(self, name):
        """Removes a defined name by name string."""
        for i, dn in enumerate(self._names):
            if dn.name == name:
                return self._names.pop(i)
        return None

    def __getitem__(self, key):
        """Gets a defined name by index or name string."""
        if isinstance(key, int):
            return self._names[key]
        elif isinstance(key, str):
            for dn in self._names:
                if dn.name == key:
                    return dn
            raise KeyError(f"Defined name '{key}' not found")
        raise TypeError("Key must be int or str")

    def __len__(self):
        return len(self._names)

    def __iter__(self):
        return iter(self._names)


class WorkbookProperties:
    """
    Container for all workbook-level properties.

    Examples:
        >>> wb.properties.file_version.app_name = "xl"
        >>> wb.properties.protection.lock_structure = True
        >>> wb.properties.view.active_tab = 0
        >>> wb.properties.calculation.calc_mode = "auto"
    """

    def __init__(self):
        self._file_version = FileVersion()
        self._workbook_pr = WorkbookPr()
        self._protection = WorkbookProtection()
        self._view = WorkbookView()
        self._calculation = CalculationProperties()
        self._defined_names = DefinedNameCollection()

    @property
    def file_version(self):
        """Gets file version properties."""
        return self._file_version

    @property
    def workbook_pr(self):
        """Gets workbook properties (workbookPr)."""
        return self._workbook_pr

    @property
    def protection(self):
        """Gets workbook protection settings."""
        return self._protection

    @property
    def view(self):
        """Gets workbook view settings."""
        return self._view

    @property
    def calculation(self):
        """Gets calculation properties."""
        return self._calculation

    @property
    def defined_names(self):
        """Gets the collection of defined names."""
        return self._defined_names
