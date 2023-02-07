from zipfile import ZipFile

import pytest

from exmlrd.sharedstyle import RichText, SharedStyle, SiTag


@pytest.fixture
def richtext():
    yield RichText()
    
@pytest.fixture
def sharedstyle():
    archive = ZipFile("tests/sample.xlsx")
    yield SharedStyle(archive)

def test_richtext():
    rich = RichText(
        sz = "16",
        color = "#FFFF",
        rFont = "メイリオ",
        family = "1",
        scheme = "",
        charset = "2",
        text = "m"
    )
    
    assert rich.sz == "16"
    assert rich.color == "#FFFF"
    assert rich.rFont == "メイリオ"
    assert rich.family == "1"
    assert rich.scheme == ""
    assert rich.charset == "2"
    assert rich.text == "m"


def test_sitag():
    rich1 = RichText(
        sz = "16",
        color = "#FFFF",
        rFont = "メイリオ",
        family = "1",
        scheme = "",
        charset = "2",
        text = "m"
    )
    
    rich2 = RichText(
        sz = "10",
        color = "#0000",
        rFont = "Symbol",
        family = "10",
        scheme = "22",
        charset = "0",
        text = "m"
    )
    
    si = SiTag()
    si.rpr.append(rich1)
    si.rpr.append(rich2)
    assert si.rpr == [rich1, rich2]

def test_si_text():
    pred = "mm"
    rich1 = RichText(
        sz = "16",
        color = "#FFFF",
        rFont = "メイリオ",
        family = "1",
        scheme = "",
        charset = "2",
        text = "m"
    )
    
    rich2 = RichText(
        sz = "10",
        color = "#0000",
        rFont = "Symbol",
        family = "10",
        scheme = "22",
        charset = "0",
        text = "m"
    )
    
    si = SiTag()
    si.rpr.append(rich1)
    si.rpr.append(rich2)
    assert si.text == pred