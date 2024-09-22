import sheet as sheet
import cell as cell
import dict_of_cells as dict_of_cells
import main


def test_calculate():
    sheet1 = sheet.Sheet('Sheet1')
    a = cell.Cell('a', '1', '1')
    b = cell.Cell('a', '2', '2')
    c = cell.Cell('a', '3', '3')
    e = cell.Cell('b', '1', '4')
    f = cell.Cell('b', '2', '5')
    g = cell.Cell('b', '3', '6')
    d = dict_of_cells.Dictionary_of_Cells('A')
    d.add_cell(a)
    d.add_cell(b)
    d.add_cell(c)
    s = dict_of_cells.Dictionary_of_Cells('B')
    s.add_cell(e)
    s.add_cell(f)
    s.add_cell(g)
    sheet1.add_column(d)
    sheet1.add_column(s)
    main.calculate(sheet1, 'A1 + A2 + A3', 'SUM', 'A4')
    assert sheet1.get_cols()['A'][3].get_value() == 6
    main.calculate(sheet1, 'B1:B3', 'SUM', 'B4')
    assert sheet1.get_cols()['B'][3].get_value() == 15
    main.calculate(sheet1, 'A1:A2, B1:B2', 'MIN', 'C1')
    assert sheet1.get_cols()['C'][0].get_value() == 1
    main.calculate(sheet1, 'A1:A3, B1:B3', 'MAX', 'C2')
    assert sheet1.get_cols()['C'][1].get_value() == 6
    main.calculate(sheet1, 'A1:A3, B1:B2', 'AVERAGE', 'C3')
    assert sheet1.get_cols()['C'][2].get_value() == 3
    assert main.calculate(sheet1, 'A1:A3 B1:B2', 'SUM', 'C3') is None
    assert main.calculate(sheet1, 'A1:A3, B1', 'SUM', 'C3') is None
    assert main.calculate(sheet1, 'A1:A3, B1:B2', 'SUM', 'C') is None
    assert main.calculate(sheet1, 'A1:A3, B1:B2', 'mult', 'C3') is None
    main.calculate(sheet1, 'A1', 'SUM', 'C3')
    assert sheet1.get_cols()['C'][2].get_value() == 1
    assert main.calculate(sheet1, 'A1:', 'SUM', 'C3') is None


def test_split_string():
    assert main.split_string('A1') == ('A', '1')
    assert main.split_string('A2') == ('A', '2')
    assert main.split_string('B3') == ('B', '3')
    assert main.split_string('B4') == ('B', '4')


def test_expand_list():
    assert ('A', 2) in main.expand_list([('A', 1), ('A', 3)])
    assert ('A', 2) in main.expand_list([('A', 1), ('A', 3), ('B', 2), ('B', 3)])


def test_convert_to_str():
    assert main.convert_to_str(['test1', 'test2']) == 'test1 test2'
    assert main.convert_to_str(['test1', 'test2', 'test3']) == 'test1 test2 test3'
    assert main.convert_to_str(['test1']) == 'test1'
