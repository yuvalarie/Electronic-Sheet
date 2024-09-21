from cell import *
import pytest
from unittest.mock import patch


def test_check_type():
    assert check_type('a') == 'a'
    assert check_type('A') == 'A'
    assert check_type('123') == 123
    assert check_type('1.23') == 1.23
    assert check_type('-456') == -456
    assert check_type('-4.56') == -4.56
    assert check_type('abcd1') == 'abcd1'
    assert check_type('1klmno') == '1klmno'
    assert check_type('2xyz3') == '2xyz3'
    assert check_type('0') == 0
    assert check_type('0.0') == 0.0
    assert check_type('-0') == 0
    assert check_type('-0.0') == 0.0
    assert check_type('= 2 + 3') == 5
    assert check_type('= 2 * 3') == 6
    assert check_type('= 2 - 3') == -1
    assert check_type('= 2 / 3') == 0.67
    with pytest.raises(ZeroDivisionError):
        check_type('= 2 / 0')
    assert check_type('= 2 // 3') == '= 2 // 3'
    assert check_type('= 2 + 3 + 4') == 9
    assert check_type('= 2 * 3 * 4') == 24
    assert check_type('= 2 - 3 - 4') == -5
    assert check_type('= 2 / 3 / 4') == 0.17
    with pytest.raises(ZeroDivisionError):
        check_type('= 2 / 0 / 4')
    with pytest.raises(ZeroDivisionError):
        check_type('= 2 / 3 / 0')
    assert check_type('= 2.5 + 3.5') == 6.0
    assert check_type('= 2.5 * 3.5') == 8.75
    assert check_type('= 2.5 - 3.5') == -1.0
    assert check_type('= 2.5 / 3.5') == 0.71
    with pytest.raises(ZeroDivisionError):
        check_type('= 2.5 / 0')
    assert check_type('= 2.5 // 3.5') == '= 2.5 // 3.5'
    assert check_type('= 2.5 + 3') == 5.5
    assert check_type('= -2.5 + 3') == 0.5
    assert check_type('= 2.5 - 3') == -0.5
    assert check_type('= 2.5 * 3') == 7.5
    assert check_type('= 2.5 / 3') == 0.83
    assert check_type('= 2.5 : 3') == '= 2.5 : 3'
    assert check_type('= 2.5 % 3') == '= 2.5 % 3'
    assert check_type('= 2.5 X 3') == '= 2.5 X 3'
    assert check_type('= 2.5 ++ 4') == '= 2.5 ++ 4'


def test_init():
    a = Cell('A','1','test')
    assert a.col == 'A'
    assert a.row == 1
    assert a.value == 'test'
    assert type(a.value) == str
    a = Cell('b', '1', '1.11')
    assert a.col == 'B'
    assert a.row == 1
    assert a.value == 1.11
    assert type(a.value) == float


@patch('builtins.input', return_value='1')
def test_get_row(capsys):
    with capsys.disabled():
        with pytest.raises(ValueError):
            a = Cell('A', 'A', 'test')
        with pytest.raises(ValueError):
            a = Cell('A', 'AB', 'test')
        with pytest.raises(ValueError):
            a = Cell('A', 'A1', 'test')
        with pytest.raises(ValueError):
            a = Cell('A', '0', 'test')
        with pytest.raises(ValueError):
            a = Cell('A', '-1', 'test')
        with pytest.raises(ValueError):
            a = Cell('A', '1.1', 'test')
        with pytest.raises(ValueError):
            a = Cell('A', '1.0', 'test')


@patch('builtins.input', return_value='A')
def test_get_col(capsys):
    with capsys.disabled():
        a = Cell('B', '1', 'test')
        assert a.col == 'B'
        a = Cell('b', '1', 'test')
        assert a.col == 'B'
        a = Cell('AB', '1', 'test')
        assert a.col == 'AB'
        a = Cell('aB', '1', 'test')
        assert a.col == 'AB'
        with pytest.raises(ValueError):
            a = Cell('a1', '1', 'test')
        with pytest.raises(ValueError):
            a = Cell('a0', '1', 'test')
        with pytest.raises(ValueError):
            a = Cell('a-1', '1', 'test')
        with pytest.raises(ValueError):
            a = Cell('a1.1', '1', 'test')
        with pytest.raises(ValueError):
            a = Cell('a1,0', '1', 'test')


def test_calculate_mathematical_expression():
    assert calculate_expression('2 + 3') == 5
    assert calculate_expression('2 * 3') == 6
    assert calculate_expression('2 - 3') == -1
    assert calculate_expression('2 / 3') == 0.6666666666666666
    with pytest.raises(ZeroDivisionError):
        calculate_expression('2 / 0')
    assert calculate_expression('2 // 3') is None
    assert calculate_expression('2 + 3 + 4') == 9
    assert calculate_expression('2 * 3 * 4') == 24
    assert calculate_expression('2 - 3 - 4') == -5
    assert calculate_expression('2 / 3 / 4') == 0.16666666666666666
    with pytest.raises(ZeroDivisionError):
        calculate_expression('2 / 0 / 4')
    with pytest.raises(ZeroDivisionError):
        calculate_expression('2 / 3 / 0')
    assert calculate_expression('2.5 + 3.5') == 6.0
    assert calculate_expression('2.5 * 3.5') == 8.75
    assert calculate_expression('2.5 - 3.5') == -1.0
    assert calculate_expression('2.5 / 3.5') == 0.7142857142857143
    with pytest.raises(ZeroDivisionError):
        calculate_expression('2.5 / 0')
    assert calculate_expression('2.5 // 3.5') is None
    assert calculate_expression('2.5 + 3') == 5.5
    assert calculate_expression('-2.5 + 3') == 0.5
    assert calculate_expression('2.5 - 3') == -0.5
    assert calculate_expression('2.5 * 3') == 7.5
    assert calculate_expression('2.5 / 3') == 0.8333333333333334
    assert calculate_expression('2.5 : 3') is None
    assert calculate_expression('2.5 % 3') is None
    assert calculate_expression('2.5 ++ 3') is None
    assert calculate_expression('2.5 ** 3') is None
    assert calculate_expression('2.5 -- 3') is None




def test_change_value():
    a = Cell('A', 1, '-2')
    a.change_value('5')
    assert a.value == 5
    assert type(a.value) == int
    a.change_value('test')
    assert a.value == 'test'
    assert type(a.value) == str
    a.change_value('5.5')
    assert a.value == 5.5
    assert type(a.value) == float
    a.change_value('=5.5 - 3')
    assert a.value == 2.5
    assert type(a.value) == float
    a.change_value('=5.5 + 3')
    assert a.value == 8.5
    assert type(a.value) == float
    a.change_value('=5.5 * 3')
    assert a.value == 16.5
    assert type(a.value) == float
    a.change_value('=5.5 / 3')
    assert a.value == 1.83
    assert type(a.value) == float
    a.change_value('=5.5 : 0')
    assert a.value == '=5.5 : 0'
    assert type(a.value) == str
    a.change_value('=5.5 // 3')
    assert a.value == '=5.5 // 3'
    assert type(a.value) == str
    a.change_value('=5.5 % 3')
    assert a.value == '=5.5 % 3'
    assert type(a.value) == str
    a.change_value('=5.5 X 3')
    assert a.value == '=5.5 X 3'
    assert type(a.value) == str
    a.change_value('=5.5 + 3 + 4')
    assert a.value == 12.5
    assert type(a.value) == float
    a.change_value('0.0')
    assert a.value == 0
    a.change_value('-0.0')
    assert a.value == 0


def test_add_related_cell():
    a = Cell('A', 1, '-2')
    b = Cell('B', 3, '0')
    a.add_related_cell(b)
    assert a.related_cells == [b]
    c = Cell('C', 2, 'test')
    a.add_related_cell(c)
    assert a.related_cells == [b, c]
    a.add_related_cell(b)
    assert a.related_cells == [b, c]
    a.add_related_cell(a)
    assert a.related_cells == [b, c, a]
    with pytest.raises(ValueError):
        a.add_related_cell('test')


def test_set_formula():
    a = Cell('A', 1, '-2')
    a.set_formula('=5.5 + 3 + 4')
    assert a.formula == '=5.5 + 3 + 4'
    a.set_formula('=5.5 + 3 * 4 + 5')
    assert a.formula == '=5.5 + 3 * 4 + 5'
    a.set_formula('=5.5 + 3 * 4 + 5 / 2')
    assert a.formula == '=5.5 + 3 * 4 + 5 / 2'
    a.set_formula('=SUM(A1 + A2 + A3 + A4 + A5)')
    assert a.formula == '=SUM(A1 + A2 + A3 + A4 + A5)'
    a.set_formula('=MIN(A1 + A2 + A3 + A4 + A5)')
    assert a.formula == '=MIN(A1 + A2 + A3 + A4 + A5)'
    a.set_formula('=MAX(A1 + A2 + A3 + A4 + A5)')
    assert a.formula == '=MAX(A1 + A2 + A3 + A4 + A5)'
    a.set_formula('=AVERAGE(A1 + A2 + A3 + A4 + A5)')
    assert a.formula == '=AVERAGE(A1 + A2 + A3 + A4 + A5)'


def test_to_dict():
    a = Cell('A', 1, '-2')
    assert a.to_dict() == {'column': 'A', 'row': 1, 'value': -2, 'formula': '', 'related_cells': []}
    a.change_value('5')
    a.set_formula('=5.5 + 3')
    assert a.to_dict() == {'column': 'A', 'row': 1, 'value': 5, 'formula': '=5.5 + 3', 'related_cells': []}
    a.change_value('test')
    assert a.to_dict() == {'column': 'A', 'row': 1, 'value': 'test', 'formula': '=5.5 + 3', 'related_cells': []}
    a.change_value('5.5')
    assert a.to_dict() == {'column': 'A', 'row': 1, 'value': 5.5, 'formula': '=5.5 + 3', 'related_cells': []}
    a.change_value('=5.5 - 3')
    assert a.to_dict() == {'column': 'A', 'row': 1, 'value': 2.5, 'formula': '=5.5 + 3', 'related_cells': []}
