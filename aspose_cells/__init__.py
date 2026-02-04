"""
Aspose.Cells for Python

A Python library for creating, reading, and modifying Excel files (.xlsx format).
This library provides a simple API compatible with Aspose.Cells for .NET structure.

Main Classes:
    - Workbook: Represents an Excel workbook
    - Worksheet: Represents a worksheet within a workbook
    - Cell: Represents a single cell
    - Cells: Collection of cells in a worksheet
    - Style: Represents cell formatting styles
"""

from .workbook import Workbook, SaveFormat
from .worksheet import Worksheet
from .cell import Cell
from .cells import Cells
from .style import Style, Font, NumberFormat
from .encryption_params import (
    AgileEncryptionParameters,
    StandardEncryptionParameters,
    CipherAlgorithm,
    HashAlgorithm,
    get_default_encryption_params
)
from .xlsx_encryptor import encrypt_xlsx, decrypt_xlsx
from .data_validation import (
    DataValidation,
    DataValidationCollection,
    DataValidationType,
    DataValidationOperator,
    DataValidationAlertStyle,
    DataValidationImeMode
)
from .csv_handler import (
    CSVHandler,
    CSVLoadOptions,
    CSVSaveOptions,
    load_csv_workbook,
    save_workbook_as_csv
)
from .markdown_handler import (
    MarkdownHandler,
    MarkdownSaveOptions,
    save_workbook_as_markdown
)
from .json_handler import (
    JsonHandler,
    JsonSaveOptions,
    save_workbook_as_json
)

__version__ = "26.2.2"
__all__ = [
    "Workbook",
    "SaveFormat",
    "Worksheet",
    "Cell",
    "Cells",
    "Style",
    "Font",
    "NumberFormat",
    "AgileEncryptionParameters",
    "StandardEncryptionParameters",
    "CipherAlgorithm",
    "HashAlgorithm",
    "get_default_encryption_params",
    "encrypt_xlsx",
    "decrypt_xlsx",
    "DataValidation",
    "DataValidationCollection",
    "DataValidationType",
    "DataValidationOperator",
    "DataValidationAlertStyle",
    "DataValidationImeMode",
    "CSVHandler",
    "CSVLoadOptions",
    "CSVSaveOptions",
    "load_csv_workbook",
    "save_workbook_as_csv",
    "MarkdownHandler",
    "MarkdownSaveOptions",
    "save_workbook_as_markdown",
    "JsonHandler",
    "JsonSaveOptions",
    "save_workbook_as_json",
]
