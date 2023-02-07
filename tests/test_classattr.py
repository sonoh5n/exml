from dataclasses import dataclass

import pytest

from exmlrd.tools import set_classattr


@dataclass
class MyTestClass:
    key1: str
    key2: str

def test_set_classattr_literal():
    data = {'key1': 'value1', 'key2': 'value2'}
    obj = MyTestClass(key1="", key2="")
    set_classattr(obj, key="key1", value=data["key1"])
    assert obj.key1 == 'value1'
    set_classattr(obj, key="key2", value=data["key2"])
    assert obj.key2 == 'value2'

def test_set_classattr_dict():
    data = {'key1': 'value1', 'key2': 'value2'}
    obj = MyTestClass(key1="", key2="")
    set_classattr(obj, container=data)
    assert obj.key1 == 'value1'
    