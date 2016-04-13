import pytest

import os.path
import sys

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../../')
sys.path.insert(0, path)

from koala.excel.utils import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
    absolute_coordinate,
    get_column_interval,
)


def test_coordinates():
    assert coordinate_from_string('ZF46') == ("ZF", 46)


@pytest.mark.parametrize("value", ['AAA', "AQ0"])
def test_invalid_coordinate(value):
    from ..exceptions import CellCoordinatesException
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string(value)


@pytest.mark.parametrize("coord, result",
                         [
                             ("ZF51",'$ZF$51',),
                             ('ZF51:ZF53', '$ZF$51:$ZF$53')
                         ]
                         )
def test_absolute(coord, result):
    assert absolute_coordinate(coord) == result


def test_column_interval():
    expected = ['A', 'B', 'C', 'D']
    assert get_column_interval('A', 'D') == expected
    assert get_column_interval('A', 4) == expected


@pytest.mark.parametrize("column, idx",
                         [
                         ('j', 10),
                         ('Jj', 270),
                         ('JJj', 7030),
                         ('A', 1),
                         ('Z', 26),
                         ('AA', 27),
                         ('AZ', 52),
                         ('BA', 53),
                         ('BZ',  78),
                         ('ZA',  677),
                         ('ZZ',  702),
                         ('AAA',  703),
                         ('AAZ',  728),
                         ('ABC',  731),
                         ('AZA', 1353),
                         ('ZZA', 18253),
                         ('ZZZ', 18278),
                         ]
                         )
def test_column_index(column, idx):
    assert column_index_from_string(column) == idx


@pytest.mark.parametrize("column",
                         ('JJJJ', '', '$', '1',)
                         )
def test_bad_column_index(column):
    with pytest.raises(ValueError):
        column_index_from_string(column)


@pytest.mark.parametrize("value", (0, 18729))
def test_column_letter_boundries(value):
    with pytest.raises(ValueError):
        get_column_letter(value)

@pytest.mark.parametrize("value, expected",
                         [
                        (18278, "ZZZ"),
                        (7030, "JJJ"),
                        (28, "AB"),
                        (27, "AA"),
                        (26, "Z")
                         ]
                         )
def test_column_letter(value, expected):
    assert get_column_letter(value) == expected


def test_coordinate_tuple():
    from .. import coordinate_to_tuple
    assert coordinate_to_tuple("D15") == (15, 4)



@pytest.mark.parametrize("range_string, sheetname, boundaries",
                         [
                        ("Sheet1!$A$1:$A$12", "Sheet1", (1,1, 1,12)),
                        ("'My Sheet'!A1:E6", "My Sheet", (1,1, 5,6)),
                         ]
                         )
def test_range_to_tuple(range_string, sheetname, boundaries):
    from .. import range_to_tuple
    assert range_to_tuple(range_string) == (sheetname, boundaries)


def test_invalid_range():
    from .. import range_to_tuple
    with pytest.raises(ValueError):
        range_to_tuple("A1:E5")


def test_quote_sheetname():
    from .. import quote_sheetname
    title = b'In D\xc3\xbcsseldorf'.decode("utf-8")
    assert quote_sheetname(title) == u"'{0}'".format(title)


def test_rows_from_range():
    from .. import rows_from_range
    cells = rows_from_range("A1:D4")
    cells = [list(row) for row in cells]
    assert cells == [
       ['A1', 'B1', 'C1', 'D1'],
       ['A2', 'B2', 'C2', 'D2'],
       ['A3', 'B3', 'C3', 'D3'],
       ['A4', 'B4', 'C4', 'D4'],
                           ]


def test_cols_from_range():
    from .. import cols_from_range
    cells = cols_from_range("A1:D4")
    cells = [list(row) for row in cells]
    assert cells == [
       ['A1', 'A2', 'A3', 'A4'],
       ['B1', 'B2', 'B3', 'B4'],
       ['C1', 'C2', 'C3', 'C4'],
       ['D1', 'D2', 'D3', 'D4'],
                           ]
