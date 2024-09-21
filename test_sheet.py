from sheet import *
from cell import *
from dict_of_cells import *
import pytest


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


def test_get_value_cols():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    z = Sheet('Y')
    z.add_column(d)
    assert z.get_value_cols() == {'A': [1, 2, 3]}


def test_to_dict():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    e = cell.Cell('b', 1, '1')
    f = cell.Cell('b', 2, '2')
    g = cell.Cell('b', 3, '3')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.to_dict() == {'A': [{'column': 'A', 'row': 1, 'value': 1, 'formula': '', 'related_cells': []}, {'column': 'A', 'row': 2, 'value': 2, 'formula': '', 'related_cells': []}, {'column': 'A', 'row': 3, 'value': 3, 'formula': '', 'related_cells': []}], 'B': [{'column': 'B', 'row': 1, 'value': 1, 'formula': '', 'related_cells': []}, {'column': 'B', 'row': 2, 'value': 2, 'formula': '', 'related_cells': []}, {'column': 'B', 'row': 3, 'value': 3, 'formula': '', 'related_cells': []}]}


def test_add_column():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    e = cell.Cell('b', 1, '4')
    f = cell.Cell('b', 2, '5')
    g = cell.Cell('b', 3, '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.get_cols() == {'A': d.get_cells(), 'B': s.get_cells()}


def test_add_row():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    e = cell.Cell('b', 1, '4')
    f = cell.Cell('b', 2, '5')
    g = cell.Cell('b', 3, '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    r = cell.Cell('A', 4, '8')
    t = cell.Cell('B', 4, '9')
    u = cell.Cell('c', 4, '10')
    w = {'A': r, 'B': t, 'C': u}
    z.add_row(w)
    assert z.get_cols() == {'A': [a, b, c, r], 'B': [e, f, g, t], 'C': z.get_cols()['C']}
    v = cell.Cell('A', '5', '8')
    n = cell.Cell('B', '4', '9')
    m = cell.Cell('c', '4', '10')
    x = {'A': v, 'B': n, 'C': m}
    z.add_row(x)
    assert z.get_cols() == {'A': [a, b, c, r, v], 'B': [e, f, g, t, n], 'C': z.get_cols()['C']}


def test_add_cell():
    z = Sheet('Y')
    a = cell.Cell('a', '1', '1')
    z.add_cell(a)
    assert z.get_cols() == {'A': [a]}
    b = cell.Cell('a', '2', '2')
    z.add_cell(b)
    assert z.get_cols() == {'A': [a, b]}
    c = cell.Cell('a', '3', '3')
    z.add_cell(c)
    assert z.get_cols() == {'A': [a, b, c]}
    d = cell.Cell('b', '1', '4')
    z.add_cell(d)
    assert d in z.get_cols()['B']
    e = cell.Cell('b', '2', '5')
    z.add_cell(e)
    assert z.get_cols()['B'][1].get_value() == e.get_value()
    f = cell.Cell('b', '3', '6')
    z.add_cell(f)
    assert z.get_value_cols() == {'A': [a.get_value(), b.get_value(), c.get_value()], 'B': [d.get_value(), e.get_value(), f.get_value()]}
    y = cell.Cell('c', '7','0')
    print(z)
    z.add_cell(y)
    print(z)
    assert z.get_value_cols()['C'][6] == y.get_value()


def test_update_cell():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    z.update_cell(z.get_cols()['A'][2], 10)
    assert z.get_cols()['A'][2].get_value() == 10
    z.update_cell(z.get_cols()['B'][2], 10)
    assert z.get_cols()['B'][2].get_value() == 10


def test_update_column():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    r = dict_of_cells('A')
    h = cell.Cell('a', '1', '4')
    i = cell.Cell('a', '2', '5')
    j = cell.Cell('a', '3', '6')
    r.add_cell(h)
    r.add_cell(j)
    r.add_cell(i)
    z.update_column(r)
    assert z.get_value_cols() == {'A': [4, 5, 6], 'B': [4, 5, 6]}


def test_update_row():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    r = cell.Cell('A', 3, '8')
    t = cell.Cell('B', 3, '9')
    w = {'A': r, 'B': t}
    z.update_row(w)
    assert z.get_value_cols() == {'A': [1, 2, 8], 'B': [4, 5, 9]}


def test_find_cell():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.find_cell(('A', 1)) == a
    assert z.find_cell(('A', 2)) == b
    assert z.find_cell(('A', 3)) == c
    assert z.find_cell(('B', 1)) == e
    assert z.find_cell(('B', 2)) == f
    assert z.find_cell(('B', 3)) == g
    assert z.find_cell(('A', 4)) is None
    assert z.find_cell(('B', 4)) is None
    assert z.find_cell(('C', 1)) is None


def test_check_if_in():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.check_if_in(('A', 1)) is True
    assert z.check_if_in(('A', 2)) is True
    assert z.check_if_in(('A', 3)) is True
    assert z.check_if_in(('B', 1)) is True
    assert z.check_if_in(('B', 2)) is True
    assert z.check_if_in(('B', 3)) is True
    assert z.check_if_in(('A', 4)) is False
    assert z.check_if_in(('B', 4)) is False
    assert z.check_if_in(('C', 1)) is False


def test_calculate_sum():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    e = cell.Cell('b', 1, '4')
    f = cell.Cell('b', 2, '5')
    g = cell.Cell('b', 3, '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.calculate_sum([('B', 1), ('B', 2), ('B', 3)], (Cell('B', '4', 'Nan'))) == 15
    assert z.calculate_sum([('A', 1), ('A', 2), ('A', 3)], (Cell('A', '4', 'Nan'))) == 6


def test_calculate_average():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    e = cell.Cell('b', 1, '4')
    f = cell.Cell('b', 2, '5')
    g = cell.Cell('b', 3, '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.calculate_average([('B', 1), ('B', 2), ('B', 3)], (Cell('B', '4', 'Nan'))) == 5.0
    assert z.calculate_average([('A', 1), ('A', 2), ('A', 3)], (Cell('A', '4', 'Nan'))) == 2.0


def test_find_minimum():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    e = cell.Cell('b', 1, '4')
    f = cell.Cell('b', 2, '5')
    g = cell.Cell('b', 3, '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.find_minimum([('B', 1), ('B', 2), ('B', 3)], (Cell('B', '4', 'Nan'))) == 4
    assert z.find_minimum([('A', 1), ('A', 2), ('A', 3)], (Cell('A', '4', 'Nan'))) == 1


def test_find_maximum():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    e = cell.Cell('b', 1, '4')
    f = cell.Cell('b', 2, '5')
    g = cell.Cell('b', 3, '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.find_maximum([('B', 1), ('B', 2), ('B', 3)], (Cell('B', '4', 'Nan'))) == 6
    assert z.find_maximum([('A', 1), ('A', 2), ('A', 3)], (Cell('A', '4', 'Nan'))) == 3


def test_calculate_cell():
    a = cell.Cell('a', 1, '1')
    b = cell.Cell('a', 2, '2')
    c = cell.Cell('a', 3, '3')
    e = cell.Cell('b', 1, '4')
    f = cell.Cell('b', 2, '5')
    g = cell.Cell('b', 3, '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    z.calculate_cell([('B', 1), ('B', 2), ('B', 3)], ('B', 4), 'SUM')
    assert z.get_cols()['B'][3].get_value() == 15
    z.calculate_cell([('A', 1), ('A', 2), ('A', 3)], ('A', 4), 'SUM')
    assert z.get_cols()['A'][3].get_value() == 6
    z.calculate_cell([('B', 1), ('B', 2), ('B', 3)], ('B', 5), 'AVERAGE')
    assert z.get_cols()['B'][4].get_value() == 5.0
    z.calculate_cell([('A', 1), ('A', 2), ('A', 3)], ('A', 5), 'MIN')
    assert z.get_cols()['A'][4].get_value() == 1
    z.calculate_cell([('B', 1), ('B', 2), ('B', 3)], ('B', 6), 'MAX')
    assert z.get_cols()['B'][5].get_value() == 6


def test_update_related_cells():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    z.calculate_cell([('B', 1), ('B', 2), ('B', 3)], ('B', 4), 'SUM')
    assert z.get_cols()['B'][3].get_value() == 15
    assert z.get_cols()['B'][3].get_formula() == '=SUM(B1 + B2 + B3)'
    assert z.get_cols()['B'][1].related_cells == [z.get_cols()['B'][3]]
    assert z.get_cols()['B'][0].related_cells == [z.get_cols()['B'][3]]
    assert z.get_cols()['B'][2].related_cells == [z.get_cols()['B'][3]]
    z.update_cell(z.get_cols()['B'][1], 10)
    assert z.get_cols()['B'][3].get_value() == 20
    z.calculate_cell([('A', 1), ('A', 2), ('A', 3)], ('A', 4), 'AVERAGE')
    assert z.get_cols()['A'][3].get_value() == 2.0
    assert z.get_cols()['A'][3].get_formula() == '=AVERAGE(A1 + A2 + A3 / 3)'
    assert z.get_cols()['A'][1].related_cells == [z.get_cols()['A'][3]]
    assert z.get_cols()['A'][0].related_cells == [z.get_cols()['A'][3]]
    assert z.get_cols()['A'][2].related_cells == [z.get_cols()['A'][3]]
    z.update_cell(z.get_cols()['A'][1], 10)
    assert z.get_cols()['A'][3].get_value() == 4.67
    z.calculate_cell([('B', 1), ('B', 2), ('B', 3)], ('B', 5), 'MIN')
    assert z.get_cols()['B'][4].get_value() == 4
    assert z.get_cols()['B'][4].get_formula() == '=MIN(B1 , B2 , B3)'
    assert z.get_cols()['B'][1].related_cells == [z.get_cols()['B'][3], z.get_cols()['B'][4]]
    assert z.get_cols()['B'][0].related_cells == [z.get_cols()['B'][3], z.get_cols()['B'][4]]
    assert z.get_cols()['B'][2].related_cells == [z.get_cols()['B'][3], z.get_cols()['B'][4]]
    z.update_cell(z.get_cols()['B'][2], 10)
    assert z.get_cols()['B'][4].get_value() == 4
    assert z.get_cols()['B'][3].get_value() == 24
    z.calculate_cell([('A', 1), ('A', 2), ('A', 3)], ('A', 5), 'MAX')
    assert z.get_cols()['A'][4].get_value() == 10
    assert z.get_cols()['A'][4].get_formula() == '=MAX(A1 , A2 , A3)'
    assert z.get_cols()['A'][1].related_cells == [z.get_cols()['A'][3], z.get_cols()['A'][4]]
    assert z.get_cols()['A'][0].related_cells == [z.get_cols()['A'][3], z.get_cols()['A'][4]]
    assert z.get_cols()['A'][2].related_cells == [z.get_cols()['A'][3], z.get_cols()['A'][4]]
    z.update_cell(z.get_cols()['A'][0], 10)
    assert z.get_cols()['A'][4].get_value() == 10
    assert z.get_cols()['A'][3].get_value() == 7.67


def test_delete_cell():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    z.delete_cell(('A', 1))
    assert z.get_cols()['A'][0].get_value() == 'None'
    z.delete_cell(('B', 2))
    assert z.get_cols()['B'][1].get_value() == 'None'
    z.delete_cell(('A', 3))
    assert z.get_cols()['A'][2].get_value() == 'None'
    z.delete_cell(('B', 1))
    assert z.get_cols()['B'][0].get_value() == 'None'


def test_delete_column():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    z.delete_column('A')
    assert z.get_value_cols()['A'] == ['None', 'None', 'None']
    z.delete_column('B')
    assert z.get_value_cols()['B'] == ['None', 'None', 'None']


def test_delete_row():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    z.delete_row(1)
    assert z.get_value_cols()['A'] == ['None', 2, 3]
    assert z.get_value_cols()['B'] == ['None', 5, 6]
    z.delete_row(2)
    assert z.get_value_cols()['A'] == ['None', 'None', 3]
    assert z.get_value_cols()['B'] == ['None', 'None', 6]
    z.delete_row(3)
    assert z.get_value_cols()['A'] == ['None', 'None', 'None']
    assert z.get_value_cols()['B'] == ['None', 'None', 'None']


def test_calculate_count_if():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '1')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', 'test')
    f = cell.Cell('b', '2', 'test')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.calculate_count_if([('B', 1), ('B', 2), ('B', 3)], Cell('B', '4', '0'), 'test') == 2
    assert z.calculate_count_if([('A', 1), ('A', 2), ('A', 3)], Cell('A', '4', '0'), '1') == 2
    assert z.calculate_count_if([('A', 1), ('A', 2), ('A', 3)], Cell('A', '4', '0'), '3') == 1
    assert z.calculate_count_if([('A', 1), ('A', 2), ('A', 3)], Cell('A', '4', '0'), '4') == 0

def test_calculate_square():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '100')
    c = cell.Cell('a', '3', '9')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '2')
    g = cell.Cell('b', '3', 'TEST')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.calculate_square(('A', 1), Cell('A', '4', '0')) == 1
    assert z.calculate_square(('A', 2), Cell('A', '4', '0')) == 10000
    assert z.calculate_square(('A', 3), Cell('A', '4', '0')) == 81
    assert z.calculate_square(('B', 1), Cell('B', '4', '0')) == 16
    assert z.calculate_square(('B', 2), Cell('B', '4', '0')) == 4
    assert z.calculate_square(('B', 3), Cell('B', '4', '0')) is None



def test_calculate_square_root():
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '100')
    c = cell.Cell('a', '3', '9')
    e = cell.Cell('b', '1', 'test')
    f = cell.Cell('b', '2', '2')
    g = cell.Cell('b', '3', '36')
    d = dict_of_cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    z = Sheet('Y')
    z.add_column(d)
    z.add_column(s)
    assert z.calculate_sqrt(('A', 1), Cell('A', '4', '0')) == 1
    assert z.calculate_sqrt(('A', 2), Cell('A', '4', '0')) == 10
    assert z.calculate_sqrt(('A', 3), Cell('A', '4', '0')) == 3
    assert z.calculate_sqrt(('B', 1), Cell('B', '4', '0')) is None
    assert z.calculate_sqrt(('B', 2), Cell('B', '4', '0')) == 1.4142135623730951
    assert z.calculate_sqrt(('B', 3), Cell('B', '4', '0')) == 6