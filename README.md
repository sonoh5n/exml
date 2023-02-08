# A Python library for easy handling of XML files

The exmlrdrd package is a simple package for reading and extracting data from Excel files. It is ideal for projects that need to extract data from Excel files without the need for a full-featured library.

This package provides a simple interface for reading and extracting data from Excel files. It supports the xlsx format and can be easily integrated into any Python project.

With exmlrd, you can easily read and extract data from Excel files without writing complicated code. Simply import the package, open the Excel file, and extract the data you need. The package also includes options to specify the sheet or range of sheets to be extracted and the ability to extract all the data in the file.

## Installation

Use pip to install exmlrd: 


```bash
pip install exmlrd
```

> Link: <https://pypi.org/project/exmlrd/>

## Usage

### Extract the Cell

With the exmlrd package, you can easily specify the row and column numbers of a cell and extract the corresponding data. This allows you to easily access specific cell information within an Excel file.

```python
import exmlrd

excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")
cell = excel_arch.get_cell(2, 3)
if cell is None:
    # Describe the processing when the type of a cell is None:
    ...

# Get Cell Address
print(f"address: {cell.address}")
# Get Cell Text
print(f"Cell Text: {cell.value}")
# Get a Cell Font
print(f"Value: {cell.style.font}")
# Get a Cell Fill
print(f"Value: {cell.style.fill}")
# Get a Cell 
print(f"Value: {cell.style.border}")
```

### Bulk extraction of cells from Excel

The exmlrd package also allows you to extract data from multiple cells at once, making it easy to retrieve large amounts of information from an Excel file in one go.

```python
import exmlrd

excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")
cell = excel_arch.get_cell(2, 3)
if cell is None:
    # Describe the processing when the type of a cell is None:
    ...

# Retrieve values from multiple cells.
# Get the name of sheet number "2"
for i in range(0, 1):
    for j in range(0, 3):
        cell = excel_arch.get_cell(i, j, worksheet=2)
        if cell is None:
            continue
        print(f"Address: {cell.address}")
        print(f"Value: {cell.value}")
```

### Retrieve properties of cells from Excel

The package's prop.decorator contains properties of the cells.

```python
import exmlrd

excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")
cell = excel_arch.get_cell(2, 3)
if cell is None:
    # Describe the processing when the type of a cell is None:
    ...

# Retrieve values from multiple cells.
# Get the name of sheet number "2"
for i in range(0, 1):
    for j in range(0, 3):
        cell = excel_arch.get_cell(i, j, worksheet=2)
        if cell is None:
            continue
        print(f"Address: {cell.address}")
        print(f"Value: {cell.value}")
        print(f"Font: {cell.style.font}")
        print(f"Border: {cell.style.border}")
        print(f"Fill: {cell.style.fill}")
        for rich in cell.shared.rpr:
            print(f"  RichText: {rich}")
        print("-" * 50)
```

### Getting the Sheet Name

To get the name of a sheet in a spreadsheet, you can use the title attribute of the sheet object in the library you are using.

```python
import exmlrd

excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")

# Get the name of sheet number "1"
# The sheet number must start with 1. index=0 is an error.
sheet_name = excel_arch.worksheet(index=1)
print(sheet_name)
```

### Retrieve addresses of all merged cells

This is how you can retrieve addresses of all merged cells in a worksheet using `exmlrd`

```python
import exmlrd

excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")

# If merged, the range of cells is returned as a string.
# If not merged, an empty string is returned.
merge_cell = excel_arch.get_mergecell(start_cell="A1", worksheet=1)
print(merge_cell)

# Obtain all merged cell locations.
merge_cells = excel_arch.get_all_mergecell(1)
print(merge_cells)
```

### Convert Json

You can also convert the information in Excel to JSON using `to_json()`

Additionally, by specifying a filename, you can output to a JSON file.

```python
excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")
# Cell(2,3) => Json
cvt_json = excel_arch.to_json(row=2, col=3)
# Cell(2,3) => sample.json
cvt_json = excel_arch.to_json(row=2, col=3, save_path="sample.json")
```

Output:

```json
{
  "Test1": [
    {
      "row": 1,
      "col": 1,
      "address": "A1",
      "value": "Sample Data",
      "formula": "",
      "shared": {
        "rpr": [
          {
            "sz": "",
            "color": "",
            "rFont": "",
            "family": "",
            "scheme": "",
            "charset": "",
            "text": "Sample Data"
          }
        ]
      },
      "style": {
        "numFmt": {
          "id": "",
          "formatCode": ""
        },
        "font": {
          "sz": "6",
          "name": "ＭＳ Ｐゴシック",
          "family": "3",
          "charset": "128",
          "scheme": "minor",
          "color": "",
          "b": false,
          "i": false,
          "u": false,
          "strike": false,
          "outline": false,
          "shadow": false,
          "condense": "",
          "extend": "",
          "vertAlign": ""
        },
        "fill": {
          "patternFill": "none",
          "fgColor": null,
          "bgColor": null
        },
        "border": {
          "left": false,
          "right": false,
          "top": false,
          "bottom": false,
          "diagonal": false
        }
      }
    }
  ]
}
```



## In Conclusion and Looking Forward

Thank you for using this project. If you have any suggestions or questions, please feel free to reach out. We hope this project has been helpful for you.

If you find this project useful, please consider sharing it with others and contributing to its development. Your support helps us continue to improve and maintain the project.

We look forward to your continued involvement and feedback. Thank you for your time and consideration.
