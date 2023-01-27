from exmlrd import excel_archiver
from exmlrd.archive import ExcelArchive


def test_excel_archiver(setup_excel):
    result = excel_archiver("tests/sample.xlsx")
    assert isinstance(result, ExcelArchive)