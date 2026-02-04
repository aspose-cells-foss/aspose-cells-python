"""
Aspose.Cells for Python - XML AutoFilter Loader Module

This module provides the AutoFilterXMLLoader class which handles loading
autofilter data from XML format according to ECMA-376 specification.

ECMA-376 Section 18.3.1.2 defines the autoFilter element structure.
"""


class AutoFilterXMLLoader:
    """
    Handles loading autofilter data from XML format for .xlsx files.

    The AutoFilterXMLLoader class is responsible for parsing the XML
    representation of autofilter settings including filter columns,
    filter values, custom filters, color filters, dynamic filters,
    top10 filters, and sort state.

    Examples:
        >>> loader = AutoFilterXMLLoader(namespaces)
        >>> loader.load_auto_filter(worksheet, worksheet_root)
    """

    def __init__(self, namespaces):
        """
        Initializes a new instance of the AutoFilterXMLLoader class.

        Args:
            namespaces: XML namespaces dictionary for parsing.
        """
        self.ns = namespaces

    def load_auto_filter(self, worksheet, worksheet_root):
        """
        Loads auto filter settings from worksheet XML according to ECMA-376 specification.

        ECMA-376 Section 18.3.1.2 defines the autoFilter element and its children:
        - autoFilter: The main auto filter element with ref attribute
        - filterColumn: Represents filter settings for a column
        - filters: Contains filter values
        - filter: Individual filter value
        - customFilters: Contains custom filter criteria
        - colorFilter: Color filter settings
        - dynamicFilter: Dynamic filter settings
        - top10: Top 10 filter settings
        - sortState: Sort state settings

        Args:
            worksheet (Worksheet): The worksheet object to load data into.
            worksheet_root: The XML root element of the worksheet.
        """
        from .auto_filter import FilterColumn

        # Find autoFilter element
        auto_filter_elem = worksheet_root.find('.//main:autoFilter', namespaces=self.ns)
        if auto_filter_elem is None:
            return

        # Get filter range
        filter_range = auto_filter_elem.get('ref')
        if filter_range:
            worksheet.auto_filter.range = filter_range

        # Load filter columns
        filter_columns = auto_filter_elem.findall('main:filterColumn', namespaces=self.ns)
        for filter_col_elem in filter_columns:
            col_id = int(filter_col_elem.get('colId'))
            filter_column = FilterColumn(col_id)

            # Load filter button visibility
            hidden_button = filter_col_elem.get('hiddenButton')
            if hidden_button == '1':
                filter_column._filter_button = False

            # Load filters (value filters)
            filters_elem = filter_col_elem.find('main:filters', namespaces=self.ns)
            if filters_elem is not None:
                for filter_elem in filters_elem.findall('main:filter', namespaces=self.ns):
                    filter_value = filter_elem.get('val')
                    if filter_value is not None:
                        filter_column.add_filter(filter_value)

            # Load custom filters
            custom_filters_elem = filter_col_elem.find('main:customFilters', namespaces=self.ns)
            if custom_filters_elem is not None:
                for custom_filter_elem in custom_filters_elem.findall('main:customFilter', namespaces=self.ns):
                    operator = custom_filter_elem.get('operator')
                    value = custom_filter_elem.get('val')
                    if operator and value is not None:
                        # Map ECMA-376 operator names to internal names
                        operator_map = {
                            'equal': 'equal',
                            'notEqual': 'notEqual',
                            'greaterThan': 'greaterThan',
                            'lessThan': 'lessThan',
                            'greaterThanOrEqual': 'greaterThanOrEqual',
                            'lessThanOrEqual': 'lessThanOrEqual',
                            'contains': 'contains',
                            'notContains': 'notContains',
                            'beginsWith': 'beginsWith',
                            'endsWith': 'endsWith'
                        }
                        operator = operator_map.get(operator, operator)
                        filter_column.add_custom_filter(operator, value)

            # Load color filter
            color_filter_elem = filter_col_elem.find('main:colorFilter', namespaces=self.ns)
            if color_filter_elem is not None:
                color = color_filter_elem.get('dxfId') or color_filter_elem.get('rgb')
                cell_color = color_filter_elem.get('cellColor', '1') == '1'
                if color:
                    filter_column.color_filter = {
                        'color': color,
                        'cell_color': cell_color
                    }

            # Load dynamic filter
            dynamic_filter_elem = filter_col_elem.find('main:dynamicFilter', namespaces=self.ns)
            if dynamic_filter_elem is not None:
                filter_type = dynamic_filter_elem.get('type')
                value = dynamic_filter_elem.get('val')
                if filter_type:
                    filter_column.dynamic_filter = {
                        'type': filter_type,
                        'value': value
                    }

            # Load top10 filter
            top10_elem = filter_col_elem.find('main:top10', namespaces=self.ns)
            if top10_elem is not None:
                top = top10_elem.get('top', '1') == '1'
                percent = top10_elem.get('percent', '0') == '1'
                val = int(top10_elem.get('val', '10'))
                filter_column.top10_filter = {
                    'top': top,
                    'percent': percent,
                    'val': val
                }

            # Add filter column to auto filter
            worksheet.auto_filter.filter_columns[col_id] = filter_column

        # Load sort state (ECMA-376 Section 18.3.1.92)
        sort_state_elem = auto_filter_elem.find('main:sortState', namespaces=self.ns)
        if sort_state_elem is not None:
            # Look for sortCondition element
            sort_condition_elem = sort_state_elem.find('main:sortCondition', namespaces=self.ns)
            if sort_condition_elem is not None:
                sort_condition_ref = sort_condition_elem.get('ref')
                descending = sort_condition_elem.get('descending', '0')
                is_descending = descending in ('1', 'true')

                # Calculate column index from sortCondition ref and filter range
                column_index = self._calculate_sort_column_index(
                    filter_range, sort_condition_ref
                )

                worksheet.auto_filter.sort_state = {
                    'column_index': column_index,
                    'descending': is_descending
                }

    def _calculate_sort_column_index(self, filter_range, sort_condition_ref):
        """
        Calculates the zero-based column index from sortCondition ref and filter range.

        Args:
            filter_range (str): The autoFilter range (e.g., "A1:D10")
            sort_condition_ref (str): The sortCondition ref (e.g., "B2:B10")

        Returns:
            int: Zero-based column index within the filter range
        """
        from .cells import Cells

        if not filter_range or not sort_condition_ref:
            return 0

        # Parse filter range to get starting column
        if ':' in filter_range:
            start_ref, _ = filter_range.split(':')
        else:
            start_ref = filter_range
        _, filter_start_col = Cells.coordinate_from_string(start_ref)

        # Parse sort condition ref to get sort column
        if ':' in sort_condition_ref:
            sort_start_ref, _ = sort_condition_ref.split(':')
        else:
            sort_start_ref = sort_condition_ref
        _, sort_col = Cells.coordinate_from_string(sort_start_ref)

        # Calculate zero-based column index
        return sort_col - filter_start_col
