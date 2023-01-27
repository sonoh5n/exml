"""
Sample code to check if the cells specified as arguments are merged in the target Excel sheet
"""

import exml

if __name__ == "__main__":

    excel_arch = exml.excel_archiver("myInputExcelFile.xlsx")

    # If merged, the range of cells is returned as a string.
    # If not merged, an empty string is returned.
    merge_cell = excel_arch.get_mergecell(start_cell="A1", worksheet=1)

    print(merge_cell)
