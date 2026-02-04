"""
Aspose.Cells for Python - Data Validation XML Saver

This module handles serialization of DataValidation objects to ECMA-376 SpreadsheetML XML.

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


# Mapping from enum values to XML attribute values
TYPE_TO_XML = {
    DataValidationType.NONE: 'none',
    DataValidationType.WHOLE_NUMBER: 'whole',
    DataValidationType.DECIMAL: 'decimal',
    DataValidationType.LIST: 'list',
    DataValidationType.DATE: 'date',
    DataValidationType.TIME: 'time',
    DataValidationType.TEXT_LENGTH: 'textLength',
    DataValidationType.CUSTOM: 'custom',
}

OPERATOR_TO_XML = {
    DataValidationOperator.BETWEEN: 'between',
    DataValidationOperator.NOT_BETWEEN: 'notBetween',
    DataValidationOperator.EQUAL: 'equal',
    DataValidationOperator.NOT_EQUAL: 'notEqual',
    DataValidationOperator.GREATER_THAN: 'greaterThan',
    DataValidationOperator.LESS_THAN: 'lessThan',
    DataValidationOperator.GREATER_THAN_OR_EQUAL: 'greaterThanOrEqual',
    DataValidationOperator.LESS_THAN_OR_EQUAL: 'lessThanOrEqual',
}

ALERT_STYLE_TO_XML = {
    DataValidationAlertStyle.STOP: 'stop',
    DataValidationAlertStyle.WARNING: 'warning',
    DataValidationAlertStyle.INFORMATION: 'information',
}

IME_MODE_TO_XML = {
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


class DataValidationXmlSaver:
    """
    Saves DataValidation objects to ECMA-376 SpreadsheetML XML format.
    """

    def __init__(self, namespace='http://schemas.openxmlformats.org/spreadsheetml/2006/main'):
        """
        Initializes the DataValidationXmlSaver.

        Args:
            namespace (str): The SpreadsheetML namespace URI.
        """
        self.namespace = namespace

    def save_data_validations(self, validations, parent_element):
        """
        Saves a DataValidationCollection to XML as a child of the parent element.

        Args:
            validations (DataValidationCollection): The validations to save.
            parent_element (Element): The parent XML element (usually worksheet).

        Returns:
            Element or None: The created dataValidations element, or None if empty.
        """
        if not validations or validations.count == 0:
            return None

        # Create dataValidations element
        ns_prefix = '{' + self.namespace + '}'
        dv_collection = ET.SubElement(parent_element, f'{ns_prefix}dataValidations')

        # Set count attribute
        dv_collection.set('count', str(validations.count))

        # Set optional attributes
        if validations.disable_prompts:
            dv_collection.set('disablePrompts', '1')

        if validations.x_window is not None:
            dv_collection.set('xWindow', str(validations.x_window))

        if validations.y_window is not None:
            dv_collection.set('yWindow', str(validations.y_window))

        # Add each validation
        for validation in validations:
            self._save_data_validation(validation, dv_collection)

        return dv_collection

    def _save_data_validation(self, validation, parent_element):
        """
        Saves a single DataValidation to XML.

        Args:
            validation (DataValidation): The validation to save.
            parent_element (Element): The parent dataValidations element.

        Returns:
            Element: The created dataValidation element.
        """
        ns_prefix = '{' + self.namespace + '}'
        dv = ET.SubElement(parent_element, f'{ns_prefix}dataValidation')

        # Required attribute: sqref
        if validation.sqref:
            dv.set('sqref', validation.sqref)

        # Type attribute (only if not default 'none')
        if validation.type != DataValidationType.NONE:
            dv.set('type', TYPE_TO_XML.get(validation.type, 'none'))

        # Operator attribute (only if type uses operators and not default 'between')
        if validation.type in (DataValidationType.WHOLE_NUMBER, DataValidationType.DECIMAL,
                               DataValidationType.DATE, DataValidationType.TIME,
                               DataValidationType.TEXT_LENGTH):
            if validation.operator != DataValidationOperator.BETWEEN:
                dv.set('operator', OPERATOR_TO_XML.get(validation.operator, 'between'))

        # Error style (only if not default 'stop')
        if validation.alert_style != DataValidationAlertStyle.STOP:
            dv.set('errorStyle', ALERT_STYLE_TO_XML.get(validation.alert_style, 'stop'))

        # IME mode (only if not default 'noControl')
        if validation.ime_mode != DataValidationImeMode.NO_CONTROL:
            dv.set('imeMode', IME_MODE_TO_XML.get(validation.ime_mode, 'noControl'))

        # Boolean attributes (only if not default)
        if validation.allow_blank:
            dv.set('allowBlank', '1')

        # Note: In ECMA-376, showDropDown="1" means HIDE the dropdown (counterintuitive)
        # So we only set this attribute if we want to HIDE the dropdown
        if not validation.show_dropdown:
            dv.set('showDropDown', '1')

        if validation.show_input_message:
            dv.set('showInputMessage', '1')

        if validation.show_error_message:
            dv.set('showErrorMessage', '1')

        # String attributes (only if set)
        if validation.error_title:
            dv.set('errorTitle', validation.error_title)

        if validation.error_message:
            dv.set('error', validation.error_message)

        if validation.input_title:
            dv.set('promptTitle', validation.input_title)

        if validation.input_message:
            dv.set('prompt', validation.input_message)

        # Formula elements
        if validation.formula1 is not None:
            formula1_elem = ET.SubElement(dv, f'{ns_prefix}formula1')
            formula1_elem.text = validation.formula1

        if validation.formula2 is not None:
            formula2_elem = ET.SubElement(dv, f'{ns_prefix}formula2')
            formula2_elem.text = validation.formula2

        return dv

    def create_data_validations_xml(self, validations):
        """
        Creates a standalone dataValidations XML element.

        Args:
            validations (DataValidationCollection): The validations to save.

        Returns:
            Element: The dataValidations element.
        """
        # Create a temporary parent
        ns_prefix = '{' + self.namespace + '}'
        temp_parent = ET.Element('temp')

        result = self.save_data_validations(validations, temp_parent)

        if result is not None:
            return result

        # Return empty element if no validations
        return ET.Element(f'{ns_prefix}dataValidations', {'count': '0'})

    def to_xml_string(self, validations, include_declaration=False):
        """
        Converts validations to an XML string.

        Args:
            validations (DataValidationCollection): The validations to convert.
            include_declaration (bool): Whether to include XML declaration.

        Returns:
            str: The XML string.
        """
        elem = self.create_data_validations_xml(validations)

        if include_declaration:
            return '<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(
                elem, encoding='unicode')
        else:
            return ET.tostring(elem, encoding='unicode')


def save_data_validations_to_worksheet_xml(validations, worksheet_element, namespace):
    """
    Convenience function to save data validations to a worksheet XML element.

    Args:
        validations (DataValidationCollection): The validations to save.
        worksheet_element (Element): The worksheet XML element.
        namespace (str): The SpreadsheetML namespace.

    Returns:
        Element or None: The created dataValidations element.
    """
    saver = DataValidationXmlSaver(namespace)
    return saver.save_data_validations(validations, worksheet_element)
