from dict_of_cells import *
from unittest.mock import patch
from cell import *
import pytest


@patch('builtins.input', return_value='A')
def test_col(capsys):
    with capsys.disabled():
        a = Dictionary_of_Cells('B')
        assert a.get_name() == 'B'
        a = Dictionary_of_Cells('b')
        assert a.get_name() == 'B'
        a = Dictionary_of_Cells('AB')
        assert a.get_name() == 'AB'
        a = Dictionary_of_Cells('aB')
        assert a.get_name() == 'AB'
        with pytest.raises(ValueError):
            a = Dictionary_of_Cells('a1')
        with pytest.raises(ValueError):
            a = Dictionary_of_Cells('a0')
        with pytest.raises(ValueError):
            a = Dictionary_of_Cells('a-1')
        with pytest.raises(ValueError):
            a = Dictionary_of_Cells('a1.1')
        with pytest.raises(ValueError):
            a = Dictionary_of_Cells('a1.0')


def test_add_cell():
    a = Cell('a', '1', '1')
    b = Cell('a', '2', '2')
    c = Cell('a', '3', '3')
    e = Cell('b', '1', '4')
    f = Cell('b', '2', '5')
    g = Cell('b', '3', '6')
    d = Dictionary_of_Cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    assert d.get_cells() == [a, b, c]
    s = Dictionary_of_Cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    assert s.get_cells() == [e, f, g]


def test_update_cell():
    a = Cell('a', '1', '1')
    b = Cell('a', '2', '2')
    c = Cell('a', '3', '3')
    e = Cell('b', '1', '4')
    f = Cell('b', '2', '5')
    g = Cell('b', '3', '6')
    d = Dictionary_of_Cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    d.update_cell('10', 1)
    assert d.get_cells()[1].get_value() == 10
    s = Dictionary_of_Cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    s.update_cell('10', 1)
    assert s.get_cells()[1].get_value() == 10