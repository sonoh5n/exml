from zipfile import ZipFile

import pytest
from pydantic import ValidationError

from exmlrd.cell import Cell, SheetXml, Worksheet
from exmlrd.exceptions import CellOutsideRange, NotFoundSheet


@pytest.fixture
def cell():
    yield Cell

@pytest.fixture
def worksheet():
    archive = ZipFile("tests/sample.xlsx")
    yield Worksheet(archive)

@pytest.fixture
def arch():
    archive = ZipFile("tests/sample.xlsx")
    yield SheetXml(archive)

class TestCell:

    def test_init_cell(slef, setup_excel, cell):
        pred_c = 0
        pred_a = ""
        assert cell.row == pred_c
        assert cell.col == pred_c
        assert cell.address == pred_a
        assert cell.value == ""
        assert cell.formula == ""
        assert cell.shared == None
        assert cell.style == None

    def test_success_cell(self, setup_excel, cell):
        pred_row = 1
        pred_col = 2
        pred_address = "A2"
        pred_value = "sample"
        pred_formula = "A1+A2+AA3/VV2"
        pred_shared = None
        pred_styles = None
        result = cell(
            row=1,
            col=2,
            address="A2",
            value = "sample",
            formula = "A1+A2+AA3/VV2")
        assert result.row == pred_row
        assert result.col == pred_col
        assert result.address == pred_address
        assert result.value == pred_value
        assert result.formula == pred_formula
        assert result.shared == pred_shared
        assert result.style == pred_styles

    def test_faild_min_row(self, setup_excel, cell):
        expected = "Invalid row number. The maximum and minimum values for the row are exceeded."
        with pytest.raises(CellOutsideRange) as exc:
            cell(row=-1, col=1)
        assert expected in str(exc.value)

    def test_faild_max_row(self, setup_excel, cell):
        expected = "Invalid row number. The maximum and minimum values for the row are exceeded."
        with pytest.raises(CellOutsideRange) as exc:
            cell(row=-1048576, col=1)
        assert expected in str(exc.value)

    def test_faild_min_col(self, setup_excel, cell):
        expected = "Invalid col number. The maximum and minimum values for the col are exceeded."
        with pytest.raises(CellOutsideRange) as exc:
            cell(row=1, col=-1)
        assert expected in str(exc.value) 

    def test_faild_max_col(self, setup_excel, cell):
        expected = "Invalid col number. The maximum and minimum values for the col are exceeded."
        with pytest.raises(CellOutsideRange) as exc:
            cell(row=1, col=-18279)
        assert expected in str(exc.value) 
        
class TestWorksheet:

    @pytest.mark.parametrize('arg, expected', [
        (2, "The sheet could not be read. (Either the index is wrong or the sheet does not exist)"),
        (10, "The sheet could not be read. (Either the index is wrong or the sheet does not exist)")])
    def test_filed_get_worksheet(self, setup_excel, worksheet, arg, expected):
        with pytest.raises(NotFoundSheet) as exc:
            ws = worksheet.get_worksheet(arg)
        assert expected in str(exc.value)
        
        
    def test_success_get_worksheet(self, setup_excel, worksheet):
        ws = worksheet.get_worksheet(1)
        assert ws == "Test1"

    @pytest.mark.parametrize('arg, expected', [
        (1, "xl/worksheets/sheet1.xml"),
        (2.99, "xl/worksheets/sheet2.xml"),
        ("1", "xl/worksheets/sheet1.xml"),
        (30.9, "xl/worksheets/sheet30.xml")])
    def test_success_get_worksheetpath(self, setup_excel, worksheet, arg, expected):
        ws_path = worksheet.get_worksheetpath(arg)
        assert ws_path == expected

    @pytest.mark.parametrize('arg, expected', [
        ("30.9", "value is not a valid integer (type=type_error.integer)"),
        ("Test1", "value is not a valid integer (type=type_error.integer)")])
    def test_failed_get_worksheetpath(self, setup_excel, worksheet, arg, expected):
        with pytest.raises(ValidationError) as exc:
            worksheet.get_worksheetpath(arg)
        assert expected in str(exc.value)

class TestMergeCell:
    
    def test_success_merge_cell(self, setup_excel, arch):
        mgcell = arch.get_mergecell("A1", 1)
        assert mgcell == ""
        
    

    
    
    