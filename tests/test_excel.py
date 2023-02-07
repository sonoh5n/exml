import pytest

from exmlrd.excel import ExcelObj
from exmlrd.exceptions import NotSupportFmt


@pytest.fixture
def excel():
    yield ExcelObj

def test_excelobj_success(setup_excel, excel):
    pred = "tests/sample.xlsx"
    result = excel(path="tests/sample.xlsx")
    assert result.path == pred

def test_excelobj_failed(setup_excel, excel):
    pred = "not support format:tests/sample.xls"
    with pytest.raises(NotSupportFmt) as ex:
        excel(path="tests/sample.xls")
    assert pred == str(ex.value)

def test_filenotfound_failed(setup_excel, excel):
    pred = "File not exist: data/sample-test.xlsx"
    with pytest.raises(FileNotFoundError) as ex:
        excel(path="data/sample-test.xlsx")
    assert pred == str(ex.value)
    

