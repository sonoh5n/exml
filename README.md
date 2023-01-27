# A Python library for easy handling of XML files

The exml package is a simple package for reading and extracting data from Excel files. It is ideal for projects that need to extract data from Excel files without the need for a full-featured library.

This package provides a simple interface for reading and extracting data from Excel files. It supports the xlsx format and can be easily integrated into any Python project.

With exml, you can easily read and extract data from Excel files without writing complicated code. Simply import the package, open the Excel file, and extract the data you need. The package also includes options to specify the sheet or range of sheets to be extracted and the ability to extract all the data in the file.

## Installation

Use pip to install exml:

```bash
pip install exml
```

## Usage

### Extract the Cell

With the exml package, you can easily specify the row and column numbers of a cell and extract the corresponding data. This allows you to easily access specific cell information within an Excel file.

```python
import exml

excel_arch = exml.excel_archiver("myInputExcelFile.xlsx")
cell, attr, prop = excel_arch.get_cell(2, 3)

# Get a Cell
cell, attr, prop = excel_arch.get_cell(2, 3)
# Get Cell Address
print(f"address: {cell.address}")
# Get Base font
print(f"Base Font: {attr.bf}")
# Get value of a cell (value of a cell != text of a cell)
print(f"Value: {attr.v}")
```

### Bulk extraction of cells from Excel

The exml package also allows you to extract data from multiple cells at once, making it easy to retrieve large amounts of information from an Excel file in one go.

```python
import exml

excel_arch = exml.excel_archiver("myInputExcelFile.xlsx")
cell, attr, prop = excel_arch.get_cell(2, 3)

# Get the name of sheet number "2"
for i in range(0, 5):
    for j in range(5):
        cell, attr, prop = excel_arch.get_cell(i, j, worksheet=2)
        print(cell)
        print(attr)
        for d in prop.decorator:
            print(d)
        print("-" * 50)
```

### Retrieve properties of cells from Excel

The package's prop.decorator contains properties of the cells.

```python
import exml

excel_arch = exml.excel_archiver("myInputExcelFile.xlsx")
cell, attr, prop = excel_arch.get_cell(2, 3)

# Get decorator
for deco in prop.decorator:
    print(f"All: {deco}")
    print(f"Text: {deco.text}")
    print(f"Font: {deco.rFont}")
```
