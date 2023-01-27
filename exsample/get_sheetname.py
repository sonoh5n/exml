"""
Sample code to get sheet name from Excel
"""

import exml

if __name__ == "__main__":

    excel_arch = exml.excel_archiver("myInputExcelFile.xlsx")
    # Get the name of sheet number "1"
    # The sheet number must start with 1. index=0 is an error.
    sheet_name = excel_arch.worksheet(index=1)
    print(sheet_name)
