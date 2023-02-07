import pytest

from exmlrd.tools import del_namespace


def test_del_namespace(setup_excel):
    n = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}scheme"
    result = del_namespace("scheme", namespace=n)
    assert result == "scheme"
    n = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}/test/scheme"
    result = del_namespace("/test/scheme", namespace=n)
    assert result == "/test/scheme"