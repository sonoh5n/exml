"""
Sample code to convert json

You can also convert the information in Excel to JSON using to_json(). 
Additionally, by specifying a filename, you can output to a JSON file.
"""

import exmlrd

if __name__ == "__main__":

    excel_arch = exmlrd.excel_archiver("myInputExcelFile.xlsx")

    # convert Json
    # Cell(2,3) => Json
    cvt_json = excel_arch.to_json(row=2, col=3)
    # Cell(2,3) => sample.json
    cvt_json = excel_arch.to_json(row=2, col=3, save_path="sample.json")