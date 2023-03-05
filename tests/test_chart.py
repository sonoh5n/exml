from zipfile import ZipFile
import pytest

from exmlrd.chart import XmlChart

@pytest.fixture
def archive():
    archive = ZipFile("tests/samle-dataset.xlsx")
    yield XmlChart(archive)

def test_get_chartsheetname(archive):
    exception = [
        "xl/charts/chart1.xml",
        "xl/charts/chart2.xml",
        "xl/charts/chart3.xml",
        "xl/charts/chart4.xml",
        "xl/charts/chart5.xml",
        "xl/charts/chart6.xml"
    ]

    chartlist = archive.chartfiles
    assert set(chartlist) == set(exception)

