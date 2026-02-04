# Aspose.Cells for Python

A lightweight Python library for creating, reading, and modifying Excel files (.xlsx format) without requiring Microsoft Excel.

[![PyPI version](https://badge.fury.io/py/aspose-cells-for-python.svg)](https://badge.fury.io/py/aspose-cells-for-python)
[![Python](https://img.shields.io/pypi/pyversions/aspose-cells-for-python.svg)](https://pypi.org/project/aspose-cells-for-python/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **Create & Edit Excel Files**: Create new workbooks or modify existing .xlsx files
- **Cell Operations**: Read/write cell values, formulas, and apply formatting
- **Styling**: Apply fonts, colors, borders, number formats, and alignment
- **Multiple Worksheets**: Add, remove, and manage worksheets
- **Data Validation**: Add dropdown lists, number ranges, and custom validation rules
- **Comments**: Add and manage cell comments with author and rich text
- **Hyperlinks**: Create links to URLs, emails, files, and internal references
- **Auto-Filters**: Apply filtering to data ranges
- **Conditional Formatting**: Apply rules-based formatting
- **CSV/JSON/Markdown Export**: Save workbooks in multiple formats
- **Encryption**: Password-protect Excel files with AES encryption
- **Workbook Protection**: Protect workbook structure and worksheets

## Installation

```bash
pip install aspose-cells-foss
```

## Quick Start

### Create a new Excel file

```python
from aspose_cells import Workbook

# Create a new workbook
workbook = Workbook()

# Get the first worksheet
worksheet = workbook.worksheets[0]

# Set cell values
worksheet.cells["A1"].put_value("Hello")
worksheet.cells["B1"].put_value("World")
worksheet.cells["A2"].put_value(42)
worksheet.cells["B2"].put_value(3.14)

# Save the workbook
workbook.save("output.xlsx")
```

### Read an existing Excel file

```python
from aspose_cells import Workbook

# Open an existing workbook
workbook = Workbook("input.xlsx")

# Access a worksheet
worksheet = workbook.worksheets[0]

# Read cell values
value = worksheet.cells["A1"].value
print(f"Cell A1 contains: {value}")
```

### Apply styling

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]
cell = worksheet.cells["A1"]

cell.put_value("Styled Text")

# Get and modify the cell style
style = cell.get_style()
style.font.is_bold = True
style.font.color = "FF0000"  # Red
style.font.size = 14
cell.set_style(style)

workbook.save("styled.xlsx")
```

### Add data validation (dropdown list)

```python
from aspose_cells import Workbook, DataValidationType

workbook = Workbook()
worksheet = workbook.worksheets[0]

# Add a dropdown list validation
validation = worksheet.data_validations.add()
validation.type = DataValidationType.LIST
validation.formula1 = '"Option1,Option2,Option3"'
validation.add_area("A1:A10")

workbook.save("validation.xlsx")
```

### Export to CSV

```python
from aspose_cells import Workbook, SaveFormat

workbook = Workbook("input.xlsx")
workbook.save("output.csv", SaveFormat.CSV)
```

### Password protection

```python
from aspose_cells import Workbook

workbook = Workbook()
worksheet = workbook.worksheets[0]
worksheet.cells["A1"].put_value("Confidential Data")

# Save with password protection
workbook.save("protected.xlsx", password="mypassword")

# Open a password-protected file
workbook2 = Workbook("protected.xlsx", password="mypassword")
```

## Requirements

- Python 3.7 or higher
- pycryptodome >= 3.15.0
- olefile >= 0.46

## Documentation

For more examples and detailed API documentation, see the [examples](https://github.com/aspose-cells-foss/aspose-cells-python/tree/main/examples) directory.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/aspose-cells-foss/aspose-cells-python/blob/main/License/Aspose_Split-License-Agreement_2026-01-26_WIP.txt) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/aspose-cells-foss/aspose-cells-python/issues)
