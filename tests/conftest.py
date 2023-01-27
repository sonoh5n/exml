import os

import openpyxl
import pytest
from openpyxl.styles import Font


@pytest.fixture
def setup_excel():
    value = [
        [8, 31, 59, 66, 23, 90],
        [69, 57, 76, 66, 23, 90],
        [17, 19, 90, 66, 23, 90],
        [27, 38, 98, 25, 56, 89],
        [10, 12, 42, 30, 26, 1],
        [12924, 32580, 55987, 2249, 91631, 1195],
        [76625, 93205, 41148, 30761, 75471, 85406]]
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Test1"
    font = Font(name='メイリオ', size=14, bold=True, italic=True, underline='single')
    for row in range(1, 5):
        for col in range(1, 6):
            sheet.cell(row=row, column=col).value = value[row-1][col-1]
    for row in sheet:
        for cell in row:
            sheet[cell.coordinate].font = font
    sheet["H1"].value = 1000
    sheet.merge_cells('H1:I1')
    workbook.save("tests/sample.xlsx")
    yield
    os.remove("tests/sample.xlsx")
    