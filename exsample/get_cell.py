"""
Sample code to get cells from Excel

If the coordinates of the specified cell are outside of the range, 
the following message will be displayed
>> [cell(x,y)] Warning: child index out of range
"""

import exmlrd

if __name__ == "__main__":

    excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")

    # Get a Cell
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
