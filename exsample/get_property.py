"""
Sample code to get property of cells from Excel

Information that can be accessed
# Cell Coordinates
# Cell Properties
# Cell Decoration

If the coordinates of the specified cell are outside of the range, 
the following message will be displayed
>> [cell(x,y)] Warning: child index out of range
"""

import exmlrd

if __name__ == "__main__":

    excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")

    # Get a Cell
    cell, attr, prop = excel_arch.get_cell(2, 3)
    # Get Cell Address
    print(f"address: {cell.address}")
    # Get Base font
    print(f"Base Font: {attr.bf}")
    # Get value of a cell (value of a cell != text of a cell)
    print(f"Value: {attr.v}")

    # Get decorator
    # The decorator can be obtained in a list object.
    # Each element has an attribute decorator from which detailed properties can be obtained.
    # There is no attribute decorator unless there are additional cell decorations.
    for deco in prop.decorator:
        print(f"All: {deco}")
        print(f"Text: {deco.text}")
        print(f"Font: {deco.rFont}")
