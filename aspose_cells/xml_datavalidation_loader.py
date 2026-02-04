"""
Aspose.Cells for Python - Data Validation XML Loader

This module handles deserialization of DataValidation objects from ECMA-376 SpreadsheetML XML.

References:
- ECMA-376 Part 4, Section 3.3.1.30 (dataValidation)
- ECMA-376 Part 4, Section 3.3.1.31 (dataValidations)
"""

import xml.etree.ElementTree as ET
from .data_validation import (
    DataValidation, DataValidationCollection,
    DataValidationType, DataValidationOperator,
    DataValidationAlertStyle, DataValidationImeMode
)


# Mapping from XML attribute values to enum values
XML_TO_TYPE = {
    'none': DataValidationType.NONE,
    'whole': DataValidationType.WHOLE_NUMBER,
    'decimal': DataValidationType.DECIMAL,
    'list': DataValidationType.LIST,
    'date': DataValidationType.DATE,
    'time': DataValidationType.TIME,
    'textLength': DataValidationType.TEXT_LENGTH,
    'custom': DataValidationType.CUSTOM,
}

XML_TO_OPERATOR = {
    'between': DataValidationOperator.BETWEEN,
    'notBetween': DataValidationOperator.NOT_BETWEEN,
    'equal': DataValidationOperator.EQUAL,
    'notEqual': DataValidationOperator.NOT_EQUAL,
    'greaterThan': DataValidationOperator.GREATER_THAN,
    'lessThan': DataValidationOperator.LESS_THAN,
    'greaterThanOrEqual': DataValidationOperator.GREATER_THAN_OR_EQUAL,
    'lessThanOrEqual': DataValidationOperator.LESS_THAN_OR_EQUAL,
}

XML_TO_ALERT_STYLE = {
    'stop': DataValidationAlertStyle.STOP,
    'warning': DataValidationAlertStyle.WARNING,
    'information': DataValidationAlertStyle.INFORMATION,
}

XML_TO_IME_MODE = {
    'noControl': DataValidationImeMode.NO_CONTROL,
    'off': DataValidationImeMode.OFF,
    'on': DataValidationImeMode.ON,
    'disabled': DataValidationImeMode.DISABLED,
    'hiragana': DataValidationImeMode.HIRAGANA,
    'fullKatakana': DataValidationImeMode.FULL_KATAKANA,
    'halfKatakana': DataValidationImeMode.HALF_KATAKANA,
    'fullAlpha': DataValidationImeMode.FULL_ALPHA,
    'halfAlpha': DataValidationImeMode.HALF_ALPHA,
    'fullHangul': DataValidationImeMode.FULL_HANGUL,
    'halfHangul': DataValidationImeMode.HALF_HANGUL,
}


class DataValidationXmlLoader:
    """
    Loads DataValidation objects from ECMA-376 SpreadsheetML XML format.
    """

    def __init__(self, namespace='http://schemas.openxmlformats.org/spreadsheetml/2006/main'):
        """
        Initializes the DataValidationXmlLoader.

        Args:
            namespace (str): The SpreadsheetML namespace URI.
        """
        self.namespace = namespace
        self.ns_prefix = '{' + namespace + '}'

    def load_data_validations(self, parent_element):
        """
        Loads a DataValidationCollection from a parent XML element.

        Args:
            parent_element (Element): The parent XML element containing dataValidations.

        Returns:
            DataValidationCollection: The loaded validations.
        """
        validations = DataValidationCollection()

        # Find dataValidations element
        dv_collection = parent_element.find(f'{self.ns_prefix}dataValidations')

        if dv_collection is None:
            # Try without namespace
            dv_collection = parent_element.find('dataValidations')

        if dv_collection is None:
            return validations

        # Load collection-level attributes
        disable_prompts = dv_collection.get('disablePrompts', '0')
        validations.disable_prompts = disable_prompts == '1' or disable_prompts.lower() == 'true'

        x_window = dv_collection.get('xWindow')
        if x_window is not None:
            validations.x_window = int(x_window)

        y_window = dv_collection.get('yWindow')
        if y_window is not None:
            validations.y_window = int(y_window)

        # Load each dataValidation element
        for dv_elem in dv_collection.findall(f'{self.ns_prefix}dataValidation'):
            validation = self._load_data_validation(dv_elem)
            validations.add_validation(validation)

        # Try without namespace if none found
        if validations.count == 0:
            for dv_elem in dv_collection.findall('dataValidation'):
                validation = self._load_data_validation(dv_elem)
                validations.add_validation(validation)

        return validations

    def _load_data_validation(self, dv_elem):
        """
        Loads a single DataValidation from an XML element.

        Args:
            dv_elem (Element): The dataValidation XML element.

        Returns:
            DataValidation: The loaded validation.
        """
        # Get sqref (required)
        sqref = dv_elem.get('sqref', '')
        validation = DataValidation(sqref)

        # Load type
        type_str = dv_elem.get('type', 'none')
        validation.type = XML_TO_TYPE.get(type_str, DataValidationType.NONE)

        # Load operator
        operator_str = dv_elem.get('operator', 'between')
        validation.operator = XML_TO_OPERATOR.get(operator_str, DataValidationOperator.BETWEEN)

        # Load error style
        error_style_str = dv_elem.get('errorStyle', 'stop')
        validation.alert_style = XML_TO_ALERT_STYLE.get(error_style_str, DataValidationAlertStyle.STOP)

        # Load IME mode
        ime_mode_str = dv_elem.get('imeMode', 'noControl')
        validation.ime_mode = XML_TO_IME_MODE.get(ime_mode_str, DataValidationImeMode.NO_CONTROL)

        # Load boolean attributes
        allow_blank = dv_elem.get('allowBlank', '0')
        validation.allow_blank = allow_blank == '1' or allow_blank.lower() == 'true'

        # Note: In ECMA-376, showDropDown="1" means HIDE the dropdown (counterintuitive)
        show_dropdown = dv_elem.get('showDropDown', '0')
        validation.show_dropdown = not (show_dropdown == '1' or show_dropdown.lower() == 'true')

        show_input = dv_elem.get('showInputMessage', '0')
        validation.show_input_message = show_input == '1' or show_input.lower() == 'true'

        show_error = dv_elem.get('showErrorMessage', '0')
        validation.show_error_message = show_error == '1' or show_error.lower() == 'true'

        # Load string attributes
        validation.error_title = dv_elem.get('errorTitle')
        validation.error_message = dv_elem.get('error')
        validation.input_title = dv_elem.get('promptTitle')
        validation.input_message = dv_elem.get('prompt')

        # Load formula elements
        formula1_elem = dv_elem.find(f'{self.ns_prefix}formula1')
        if formula1_elem is None:
            formula1_elem = dv_elem.find('formula1')
        if formula1_elem is not None and formula1_elem.text:
            validation.formula1 = formula1_elem.text

        formula2_elem = dv_elem.find(f'{self.ns_prefix}formula2')
        if formula2_elem is None:
            formula2_elem = dv_elem.find('formula2')
        if formula2_elem is not None and formula2_elem.text:
            validation.formula2 = formula2_elem.text

        return validation

    def load_from_xml_string(self, xml_string):
        """
        Loads validations from an XML string.

        Args:
            xml_string (str): The XML string containing dataValidations.

        Returns:
            DataValidationCollection: The loaded validations.
        """
        # Parse the XML
        root = ET.fromstring(xml_string)

        # Check if root is dataValidations
        if root.tag.endswith('dataValidations'):
            # Create a temporary parent
            temp_parent = ET.Element('temp')
            temp_parent.append(root)
            return self.load_data_validations(temp_parent)

        return self.load_data_validations(root)

    def load_from_file(self, file_path):
        """
        Loads validations from an XML file.

        Args:
            file_path (str): Path to the XML file.

        Returns:
            DataValidationCollection: The loaded validations.
        """
        tree = ET.parse(file_path)
        root = tree.getroot()
        return self.load_data_validations(root)


def load_data_validations_from_worksheet_xml(worksheet_element, namespace):
    """
    Convenience function to load data validations from a worksheet XML element.

    Args:
        worksheet_element (Element): The worksheet XML element.
        namespace (str): The SpreadsheetML namespace.

    Returns:
        DataValidationCollection: The loaded validations.
    """
    loader = DataValidationXmlLoader(namespace)
    return loader.load_data_validations(worksheet_element)
