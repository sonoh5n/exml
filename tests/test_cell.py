import pytest

from exml.cell import Coordinate
from exml.exceptions import CellOutsideRange


@pytest.fixture
def coordinate():
    yield Coordinate

def test_init_coordinate(coordinate):
    pred_c = 0
    pred_a = ""
    assert coordinate.row == pred_c
    assert coordinate.col == pred_c
    assert coordinate.address == pred_a

def test_success_coordinate(coordinate):
    pred_row = 1
    pred_col = 2
    pred_address = "A2"
    result = coordinate(row=1, col=2, address=pred_address)
    assert result.row == pred_row
    assert result.col == pred_col
    assert result.address == pred_address

def test_faild_min_row(coordinate):
    expected = "Invalid row number. The maximum and minimum values for the row are exceeded."
    with pytest.raises(CellOutsideRange) as exc:
        coordinate(row=-1, col=1)
    assert expected in str(exc.value)

def test_faild_max_row(coordinate):
    expected = "Invalid row number. The maximum and minimum values for the row are exceeded."
    with pytest.raises(CellOutsideRange) as exc:
        coordinate(row=-1048576, col=1)
    assert expected in str(exc.value)

def test_faild_min_col(coordinate):
    expected = "Invalid col number. The maximum and minimum values for the col are exceeded."
    with pytest.raises(CellOutsideRange) as exc:
        coordinate(row=1, col=-1)
    assert expected in str(exc.value) 

def test_faild_max_col(coordinate):
    expected = "Invalid col number. The maximum and minimum values for the col are exceeded."
    with pytest.raises(CellOutsideRange) as exc:
        coordinate(row=1, col=-18279)
    assert expected in str(exc.value) 
    
    
    
    
    