import typing
import pandas as pd
from typing import List, Tuple, Dict, Any, Union
from dict_of_cells import Dictionary_of_Cells as dict_of_cells
from cell import Cell
import re
import numpy as np
import math


def find_cells(user_input: str) -> List[str]:
    """
       This function gets a string and returns a list of cells in the string.
       :param user_input: The string to find the cells in.
       :return: The list of cells in the string.
       """
    pattern = r'[A-Za-z]+[0-9]+'
    return re.findall(pattern, user_input)


def replace_cells_with_values(expression: str, cells: List[str], cell_values: Dict[str, Any]) -> str:
    """
        This function gets an expression, a list of cells and a dictionary of cell values and returns the expression with the
        :param expression: string that represents a mathematical expression.
        :param cells: list of cells in the expression.
        :param cell_values: dictionary of cell values.
        :return: a string that represents a mathematical expression with the values of the cells.
        """
    for cell in cells:
        if cell in cell_values:
            expression = expression.replace(cell, str(cell_values[cell]))
    return expression


def calculate_expression(expression: str) -> Any:
    """
    This function gets an expression and returns the result of the expression.
    :param expression: string that represents a mathematical expression.
    :return: the result of the expression.
    """
    # Define the allowed operators
    allowed_operators = set('+-*/. ')
    # Check for division by zero
    if '/0' in expression:
        print("Error: Division by zero is not allowed.")
        return None
    elif '//' in expression:
        return None
    elif '++' in expression:
        return None
    elif '**' in expression:
        return None
    elif '--' in expression:
        return None
    # Check for illegal operators
    if any(char not in allowed_operators for char in expression if
           not char.isdigit()):
        print("Error: Illegal operator used.")
        return None
    try:
        # Evaluate the expression
        result = eval(expression)
    except SyntaxError:
        print("Error: Invalid syntax.")
        return None
    return result


class Sheet:

    """
    This class is a sheet of cells.
    attributes:
    name: str
    cols: dict[str: list[Cell]]
    value_cols: dict[str: list[Cell]]
    data: pd.DataFrame's
    methods:
    __init__: The constructor of the class.
    __str__: The string representation of the class.
    to_dict: returns the sheet as a dictionary.
    from_dict: returns a sheet from a dictionary.
    get_cols: returns the columns of the sheet.
    get_data: returns the data of the sheet.
    get_value_cols: returns the value columns of the sheet.
    add_column: adds a column to the data.
    add_row: adds a row to the data.
    add_cell: adds a cell to the data.
    update_cell: updates a cell in the data.
    update_column: updates a column in the data.
    update_row: updates a row in the data.
    find_cell: finds a cell in the data.
    formulate_cell: formulates a cell in the data.
    update_related_cells: updates the related cells of a cell.
    reformulate_sum: recalculates the sum of a list of indexes.
    recalculate_average: recalculates the average of a list of indexes.
    refind_minimum: refinds the minimum value in a list of indexes.
    refind_maximum: refinds the maximum value in a list of indexes.
    recalculate square: recalculates the square of a cell.
    recalculate sqrt: recalculates the square root of a cell.
    recalculate_count_if: recalculates the count of cells that meet the criteria.
    calculate_sum: calculates the sum of a list of indexes.
    calculate_average: calculates the average of a list of indexes.
    find_minimum: finds the minimum value in a list of indexes.
    find_maximum: finds the maximum value in a list of indexes.
    calculate square: calculates the square of a cell.
    calculate sqrt: calculates the square root of a cell.
    calculate_count_if: calculates the count of cells that meet the criteria.
    check_if_in: checks if a cell is in the data.
    calculate_cell: calculates a cell in the data.
    delete cell: deletes a cell from the data.
    delete column: deletes a column from the data.
    delete row: deletes a row from the data.
    """

    def __init__(self, name: str) -> None:
        """
        This function is constructing a new sheet.
        :param name: The name of the sheet.
        """
        self.name = name
        self.cols: Dict[str, list[Cell]] = {}
        self.value_cols: Dict[str, list[Union[str, int, float, None]]] = {}
        self.data = pd.DataFrame(self.value_cols)

    def __str__(self) -> str:
        """
        This function creates a string representation of the sheet.
        :return: String
        """
        df_temp = self.data.copy()
        df_temp.index += 1
        df_styled = df_temp.style.set_properties(**{'font-weight': 'bold'},
                                                 subset=pd.IndexSlice[
                                                        df_temp.index, :])
        df_styled.set_properties(**{'font-weight': 'bold'},
                                 subset=pd.IndexSlice[:, df_temp.columns])
        return df_temp.to_string()

    def to_dict(self) -> Any:
        """
        This function returns the sheet as a dictionary.
        :return: A dictionary of the sheet.
        """
        cells_dict = {}
        for col in self.cols:
            cells_dict[col] = [cell.to_dict() for cell in self.cols[col]]
        return cells_dict

    @classmethod
    def from_dict(cls, data: Dict[str, Dict[str, typing.Any]]) -> 'Sheet':
        """
        This function gets a dictionary and returns a sheet from it.
        :param data: A dictionary of cells.
        :return: A sheet.
        """
        sheet = Sheet('Sheet')
        for col in data:
            temp_dict = dict_of_cells(col)
            for c in data[col]:
                temp_cell = Cell.from_dict(c)
                temp_dict.add_cell(temp_cell)
            sheet.add_column(temp_dict)
        return sheet

    def get_cols(self) -> Dict[str, list[Cell]]:
        """
        This function returns the columns of the sheet.
        :return: A dictionary of columns.
        """
        return self.cols

    def get_data(self) -> pd.DataFrame:
        """
        This function returns the data of the sheet.
        :return: A data frame.
        """
        return self.data

    def get_value_cols(self) -> Dict[str, list[Cell]]:
        return self.value_cols

    def add_column(self, col: dict_of_cells) -> None:
        """
        This function gets a column and adds it to the data.
        :param col: a column as a dictionary of cells.
        :return: None
        """
        if col.get_name() in self.cols:  # checks if the column is already in the data.
            self.update_column(col)
            return
        if self.data.empty:  # checks if the data is empty.
            self.cols[col.get_name()] = col.get_cells()
            self.value_cols[col.get_name()] = [i.get_value() for i in
                                               col.get_cells()]
            self.data = pd.DataFrame(self.value_cols)
            return
        num_of_rows = len(self.cols[self.data.columns[0]])
        if len(col.get_cells()) > num_of_rows:  # checks if the last rows is needed to be added to the exisiting cols.
            for r in range(num_of_rows, len(col.get_cells())):
                row_dict = {}
                for c in self.cols:
                    row_dict[c] = Cell(c, str(r), np.NAN)
                self.add_row(row_dict)
        elif len(col.get_cells()) < num_of_rows:  # checks if the last rows is needed to be added to the new col.
            for r in range(len(col.get_cells()), num_of_rows):
                col.add_cell(Cell(col.get_name(), str(r + 1), None))
        self.cols[col.get_name()] = col.get_cells()
        self.value_cols[col.get_name()] = [i.get_value() for i in
                                           col.get_cells()]
        self.data = pd.DataFrame(self.value_cols)  # updates the data

    def add_row(self, rows: Dict[str, 'Cell']) -> None:
        """
        This function gets a row and adds it to the data.
        :param rows: the new row that is added.
        :return: None
        """
        for col in rows:
            if col in self.cols:
                self.cols[col].append(rows[col])
                self.value_cols[col].append(rows[col].get_value())
            else:
                temp_dict = dict_of_cells(col)
                for i in range(ord(col) - 64 + 1):
                    temp_cell = Cell(col, str(i + 1), np.NAN)
                    temp_dict.add_cell(temp_cell)
                temp_dict.update_cell(rows[col].get_value(),
                                      rows[col].get_row() - 1)
                self.add_column(temp_dict)
                self.value_cols[col] = [i.get_value() for i in
                                        temp_dict.get_cells()]
        self.data = pd.DataFrame(self.value_cols)  # updates the data

    def add_cell(self, cell: Cell) -> None:
        """
        This function adds a new single cell to the data.
        :param cell: The new cell that is added.
        :return: None.
        """
        col = cell.get_col()
        row = cell.get_row()
        if self.data.empty:  # checks if the data is empty.
            temp_dict = dict_of_cells(col)
            temp_dict.add_cell(cell)
            self.add_column(temp_dict)
            return
        if self.check_if_in((col, row)):
            self.update_cell(self.get_cols()[col][row - 1], cell.get_value())
            return
        sheet_rows = self.data.index[-1]  # The last number of row in the data.
        #sheet_cols = ord(self.data.columns[-1])  # The last letter of column in the data.
        if col not in self.cols:  # checks if a new column is needed to be added.
            temp_col = col
            temp_dict = dict_of_cells(temp_col)
            for i in range(max(sheet_rows + 1, row)):
                if i + 1 == row:
                    temp_dict.add_cell(cell)
                else:
                    temp_cell = Cell(temp_col, str(i + 1), np.NAN)
                    temp_dict.add_cell(temp_cell)
            if row - 1 > sheet_rows: #checks if the last rows is needed to be added.
                for r in range(sheet_rows + 1, row):
                    row_dict = {}
                    for c in self.cols:
                       row_dict[c] = Cell(c, str(r), np.NAN)
                    self.add_row(row_dict)
            self.add_column(temp_dict)
        else:  # if a new row is needed to be added.
            for c in self.cols:
                if c == col:
                    self.cols[c].append(cell)
                    self.value_cols[c].append(cell.get_value())
                else:
                    self.cols[c].append(Cell(c, str(row), np.NAN))
                    self.value_cols[c].append(np.NAN)
            self.data = pd.DataFrame(self.value_cols)

    def update_cell(self, cell: Cell, value: str) -> None:
        """
        This function gets a single existing cell and updates it.
        :param value: the new value of the cell.
        :param cell: The new cell that needs to be updated.
        :return: None.
        """
        cell.change_value(value)
        self.value_cols[cell.get_col()][cell.get_row() - 1] = cell.get_value()
        cell.set_formula('')
        if cell.related_cells:
            self.update_related_cells(cell)
        self.data = pd.DataFrame(self.value_cols)

    def update_column(self, col: dict_of_cells) -> None:
        """
        This function gets a column and updates it.
        :param col: the new column that needs to be updated.
        :return: None.
        """
        for c in col.get_cells():
            self.update_cell(c, c.get_value())

    def update_row(self, rows: Dict[str, Cell]) -> None:
        """
        This function gets a row and updates it.
        :param rows: the new row that needs to be updated.
        :return: None.
        """
        for r in rows:
            self.update_cell(rows[r], rows[r].get_value())

    def delete_cell(self, location: Tuple[str, int]) -> None:
        """
        This function gets a location and deletes the cell in it.
        :param location: the location of the cell.
        :return: None.
        """
        if location[0] not in self.cols:
            raise ValueError
        if location[1] - 1 > self.data.index[-1]:
            raise ValueError
        cur_cell = self.cols[location[0]][location[1] - 1]
        self.update_cell(cur_cell, None)

    def delete_column(self, col: str) -> None:
        """
        This function gets a column and deletes it.
        :param col: the column that needs to be deleted.
        :return: None.
        """
        if col not in self.cols:
            raise ValueError
        for c in self.cols[col]:
            self.update_cell(c, None)

    def delete_row(self, row: int) -> None:
        """
        This function gets a row and deletes it.
        :param row: the row that needs to be deleted.
        :return: None.
        """
        for c in self.cols:
            if row - 1 > len(self.cols[c]):
                raise ValueError
            cur_cell = self.cols[c][row - 1]
            self.update_cell(cur_cell, None)

    def find_cell(self, location: Tuple[str, int]) -> Cell or None:
        """
        This function gets a location and returns the cell in it.
        :param location: the location of the cell.
        :return: The cell in the location.
        """
        if location[0] not in self.cols:
            return None
        if location[1] - 1 > self.data.index[-1]:
            return None
        return self.cols[location[0]][location[1] - 1]

    def formulate_cell(self, cell: Cell) -> None:
        """
        This function gets a single existing cell and updates it.
        :param cell: The new cell that needs to be updated.
        :return: None.
        """
        if cell.get_formula() is None or cell.get_formula() == "":
            return
        form = cell.get_formula()[1:]
        if form[:3] == 'SUM':
            form = form[4:-1]
            indexes = form.split(' + ')
            indexes = [(i[0], int(i[1])) for i in indexes]
            val = self.recalculate_sum(indexes, cell)
            cell.change_value(val)
            self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = val
        elif form[:7] == 'AVERAGE':
            form = form[8:-1]
            indexes = form.split(' + ')
            indexes = [(i[0], int(i[1])) for i in indexes]
            val = self.recalculate_average(indexes, cell)
            cell.change_value(val)
            self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = val
        elif form[:3] == 'MAX':
            form = form[4:-1]
            indexes = form.split(' , ')
            indexes = [(i[0], int(i[1])) for i in indexes]
            val = self.refind_maximum(indexes, cell)
            cell.change_value(val)
            self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = val
        elif form[:3] == 'MIN':
            form = form[4:-1]
            indexes = form.split(' , ')
            indexes = [(i[0], int(i[1])) for i in indexes]
            val = self.refind_minimum(indexes, cell)
            cell.change_value(val)
            self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = val
        elif form[:7] == 'COUNTIF':
            criteria = ""
            i = len(form) - 1
            while form[i] != ' ':  # find the criteria
                criteria += form[i]
                i -= 1
            criteria = criteria[::-1]
            form = form[8:-len(criteria) - 1]
            indexes = form.split(' , ')
            indexes = [(i[0], int(i[1])) for i in indexes]
            val = self.recalculate_count_if(indexes, cell, criteria)
            cell.change_value(str(val))
            self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = val
        elif form[:6] == 'SQUARE':
            form = form[7:-1]
            index = (form[0], int(form[1]))
            val = self.recalculate_square(index, cell)
            cell.change_value(val)
            self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = val
        elif form[:4] == 'SQRT':
            form = form[5:-1]
            index = (form[0], int(form[1]))
            val = self.recalculate_sqrt(index, cell)
            cell.change_value(val)
            self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = val
        else:
            try:
                cells = find_cells(form)
                cell_values = {}
                for c in cells:
                    location = (c[0].upper(), int(c[1]))
                    if not self.check_if_in(location):
                        print("Invalid index.")
                        raise ValueError
                    else:
                        cur_cell = self.find_cell(location)
                        cell_values[c] = cur_cell.get_value()
                try:
                    num_expression = replace_cells_with_values(form,
                                                               cells,
                                                               cell_values)
                    value = calculate_expression(num_expression)
                except:
                    print("Invalid expression.")
                    raise ValueError
                cell.change_value(value)
                self.get_value_cols()[cell.get_col()][cell.get_row() - 1] = value
            except:
                print("Invalid formula.")
                raise ValueError
            return

    def update_related_cells(self, cell: Cell) -> None:
        """
        This function updates the related cells of a cell.
        :param cell: the cell that its related cells need to be updated.
        :return: None.
        """
        self.formulate_cell(cell)
        if not cell.related_cells:
            return
        for c in cell.related_cells:
            self.update_related_cells(c)
        self.formulate_cell(cell)

    # This function is for recalculation using the formula, and not for the
    # first calculation
    def recalculate_sum(self, indexes: List[Tuple[str, int]],
                        target_cell: Cell) -> Union[int, float]:
        """
        This function recalculates the sum of a list of indexes.
        :param target_cell: The targeted cell for the sum.
        :param indexes: List of indexes.
        :return: sum.
        """
        s = 0
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " + "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            val = temp_cell.get_value()
            if val is np.NAN or val is None or type(val) == str:
                continue
            s += val
        target_cell.set_formula("=SUM(" + form[:-3] + ")")
        return s

    # This function is for recalculation using the formula, and not for the
    # first calculation
    def recalculate_average(self, indexes: List[Tuple[str, int]],
                            target_cell: Cell) -> Union[int, float]:
        """
        This function recalculates the average of a list of indexes.
        :param target_cell: the targeted cell for the average.
        :param indexes: list of indexes.
        :return: average.
        """
        s = 0
        c = 0
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " + "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            val = temp_cell.get_value()
            if val is np.NAN or val is None or type(val) == str:
                continue
            s += val
            c += 1
        target_cell.set_formula("=AVERAGE(" + form[:-3] + " / " + str(c) + ")")
        return s / c

    # This function is for recalculation using the formula, and not for the
    # first calculation
    def refind_minimum(self, indexes: List[Tuple[str, int]],
                       target_cell: Cell) -> Union[int, float, None]:
        """
        This function refinds the minimum value in a list of indexes.
        :param target_cell: the targeted cell for the minimum.
        :param indexes: a list of indexes.
        :return: minimum value.
        """
        min = math.inf
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " , "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            val = temp_cell.get_value()
            if val == np.NAN or val is None or type(val) == str:
                continue
            if val < min:
                min = val
        target_cell.set_formula("=MIN(" + form[:-3] + ")")
        if min == math.inf or isinstance(min, str):
            return None
        return min

    # This function is for recalculation using the formula, and not for the
    # first calculation
    def refind_maximum(self, indexes: List[Tuple[str, int]],
                       target_cell: Cell) -> Union[int, float, None]:
        """
        This function refinds the maximum value in a list of indexes.
        :param target_cell: the targeted cell for the maximum.
        :param indexes: a list of indexes.
        :return: maximum value.
        """
        max = -1 * math.inf
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " , "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            val = temp_cell.get_value()
            if val == np.NAN or val is None or type(val) == str:
                continue
            if val > max:
                max = val
        target_cell.set_formula("=MAX(" + form[:-3] + ")")
        if max == -1 * math.inf or isinstance(max, str):
            return None
        return max

    # This function is for recalculation using the formula, and not for the
    # first calculation
    def recalculate_count_if(self, indexes: List[Tuple[str, int]], target_cell: Cell, criteria: str) -> int:
        """
        This function recalculates the count of cells that meet the criteria.
        :param indexes: A list of indexes of the cells that the calculation will be on.
        :param target_cell: An index of the cell that the result will be put in it.
        :param criteria: The criteria for the count.
        :return: The count of cells that meet the criteria.
        """
        count = 0
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " , "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            val = temp_cell.get_value()
            if str(val) == criteria:
                count += 1
        target_cell.set_formula("=COUNTIF(" + form[:-3] + ") " + str(criteria))
        return count

    # This function is for recalculation using the formula, and not for the
    # first calculation
    def recalculate_square(self, index: Tuple[str, int], target_cell: Cell) -> Union[int, float, None]:
        """
        This function recalculates the square of a cell.
        :param index: the cell that the calculation will be on.
        :param target_cell: An index of the cell that the result will be put in it.
        :return: The square of the cell.
        """
        form = index[0] + str(index[1])
        if index[0] not in self.cols or index[1] > len(self.cols[index[0]]):
            return None
        temp_cell = self.cols[index[0]][index[1] - 1]
        val = temp_cell.get_value()
        if val == np.NAN or val is None or type(val) == str:
            return None
        target_cell.set_formula("=SQUARE(" + form + ")")
        return val ** 2

    # This function is for recalculation using the formula, and not for the
    # first calculation
    def recalculate_sqrt(self, index: Tuple[str, int], target_cell: Cell) -> Union[int, float, None]:
        """
        This function recalculates the square root of a cell.
        :param index: the cell that the calculation will be on.
        :param target_cell: An index of the cell that the result will be put in it.
        :return: The square root of the cell.
        """
        form = index[0] + str(index[1])
        if index[0] not in self.cols or index[1] > len(self.cols[index[0]]):
            return None
        temp_cell = self.cols[index[0]][index[1] - 1]
        val = temp_cell.get_value()
        if val == np.NAN or val is None or type(val) == str:
            return None
        target_cell.set_formula("=SQRT(" + form + ")")
        return math.sqrt(val)

    def calculate_sum(self, indexes: List[Tuple[str, int]],
                      target_cell: Cell) -> Union[int, float]:
        """
        This function calculates the sum of a list of indexes.
        :param target_cell: The targeted cell for the sum.
        :param indexes: List of indexes.
        :return: sum.
        """
        s = 0
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " + "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            temp_cell.add_related_cell(target_cell)
            val = temp_cell.get_value()
            if val is np.NAN or val is None or type(val) == str:
                continue
            s += val
        target_cell.set_formula("=SUM(" + form[:-3] + ")")
        return s

    def calculate_average(self, indexes: List[Tuple[str, int]],
                          target_cell: Cell) -> Union[int, float]:
        """
        This function calculates the average of a list of indexes.
        :param target_cell: the targeted cell for the average.
        :param indexes: list of indexes.
        :return: average.
        """
        s = 0
        c = 0
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " + "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            temp_cell.add_related_cell(target_cell)
            val = temp_cell.get_value()
            if val is np.NAN or val is None or type(val) == str:
                continue
            s += val
            c += 1
        target_cell.set_formula("=AVERAGE(" + form[:-3] + " / " + str(c) + ")")
        return s / c

    def find_minimum(self, indexes: List[Tuple[str, int]],
                         target_cell: Cell) ->  Union[int, float, None]:
        """
        This function finds the minimum value in a list of indexes.
        :param target_cell: the targeted cell for the minimum.
        :param indexes: a list of indexes.
        :return: minimum value.
        """
        min = math.inf
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " , "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            temp_cell.add_related_cell(target_cell)
            val = temp_cell.get_value()
            if val == np.NAN or val is None or type(val) == str:
                continue
            if val < min:
                min = val
        target_cell.set_formula("=MIN(" + form[:-3] + ")")
        if min == math.inf or isinstance(min, str):
            return None
        return min

    def find_maximum(self, indexes: List[Tuple[str, int]],
                     target_cell: Cell) -> Union[int, float, None]:
        """
        This function finds the maximum value in a list of indexes.
        :param target_cell: the targeted cell for the maximum.
        :param indexes: a list of indexes.
        :return: maximum value.
        """
        max = -1 * math.inf
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " , "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            temp_cell.add_related_cell(target_cell)
            val = temp_cell.get_value()
            if val == np.NAN or val is None or type(val) == str:
                continue
            if val > max:
                max = val
        target_cell.set_formula("=MAX(" + form[:-3] + ")")
        if max == -1 * math.inf or isinstance(max, str):
            return None
        return max

    def calculate_count_if(self, indexes: List[Tuple[str, int]], target_cell: Cell, criteria: str) -> int:
        """
        This function calculates the count of cells that meet the criteria.
        :param indexes: A list of indexes of the cells that the calculation will be on.
        :param target_cell: An index of the cell that the result will be put in it.
        :param criteria: The criteria for the count.
        :return: The count of cells that meet the criteria.
        """
        count = 0
        form = ""
        for item in indexes:
            form += item[0] + str(item[1]) + " , "
            if item[0] not in self.cols or item[1] > len(self.cols[item[0]]):
                continue
            temp_cell = self.cols[item[0]][item[1] - 1]
            temp_cell.add_related_cell(target_cell)
            val = temp_cell.get_value()
            if str(val) == criteria:
                count += 1
        target_cell.set_formula("=COUNTIF(" + form[:-3] + ") " + str(criteria))
        return count

    def calculate_square(self, index: Tuple[str, int], target_cell: Cell) -> Union[int, float, None]:
        """
        This function calculates the square of a cell.
        :param indexes: the cell that the calculation will be on.
        :param target_cell: An index of the cell that the result will be put in it.
        :return: The square of the cell.
        """
        form = index[0] + str(index[1])
        if index[0] not in self.cols or index[1] > len(self.cols[index[0]]):
            return None
        temp_cell = self.cols[index[0]][index[1] - 1]
        temp_cell.add_related_cell(target_cell)
        val = temp_cell.get_value()
        if val == np.NAN or val is None or type(val) == str:
            return None
        target_cell.set_formula("=SQUARE(" + form + ")")
        return val ** 2

    def calculate_sqrt(self, index: Tuple[str, int], target_cell: Cell) -> Union[int, float, None]:
        """
        This function calculates the square root of a cell.
        :param indexes: the cell that the calculation will be on.
        :param target_cell: An index of the cell that the result will be put in it.
        :return: The square root of the cell.
        """
        form = index[0] + str(index[1])
        if index[0] not in self.cols or index[1] > len(self.cols[index[0]]):
            return None
        temp_cell = self.cols[index[0]][index[1] - 1]
        temp_cell.add_related_cell(target_cell)
        val = temp_cell.get_value()
        if val == np.NAN or val is None or type(val) == str:
            return None
        target_cell.set_formula("=SQRT(" + form + ")")
        return math.sqrt(val)

    def check_if_in(self, location: Tuple[str, int]) -> bool:
        """
        This function checks if a cell is in the data.
        :param location: the location of the cell.
        :return: None.
        """
        if location[0] not in self.cols:
            return False
        if location[1] - 1 > self.data.index[-1]:
            return False
        return True

    def calculate_cell(self, indexes: List[Tuple[str, int]],
                       target_location: Tuple[str, int], func: str, criteria=None) -> None:
        """
        This function operates a function (SUM, AVERAGE, MAXIMUM, MINIMUM) and enters the result in the target cell.
        :param indexes: A list of indexes of the cells that the calculation will be on.
        :param target_location: An index of the cell that the result will be put in it.
        :param func: SUM, AVERAGE, MAXIMUM or MINIMUM
        :param criteria: The criteria for the COUNTIF function.
        :return: True if succeeded, False otherwise.
        """
        if target_location in indexes:
            print("The target cell can't be in the indexes.")
            return
        if not self.check_if_in(target_location):
            target_cell = Cell(target_location[0], str(target_location[1]), None)
            self.add_cell(target_cell)
        else:
            target_cell = self.cols[target_location[0]][target_location[1] - 1]
        if func == 'SUM':
            sum = self.calculate_sum(indexes, target_cell)
            target_cell.change_value(str(sum))
            self.value_cols[target_cell.get_col()][target_cell.get_row() - 1] = sum
        elif func == 'AVERAGE':
            avg = self.calculate_average(indexes, target_cell)
            target_cell.change_value(str(avg))
            self.value_cols[target_cell.get_col()][target_cell.get_row() - 1] = avg
        elif func == 'MAX':
            max = self.find_maximum(indexes, target_cell)
            target_cell.change_value(str(max))
            self.value_cols[target_cell.get_col()][target_cell.get_row() - 1] = max
        elif func == 'MIN':
            min = self.find_minimum(indexes, target_cell)
            target_cell.change_value(str(min))
            self.value_cols[target_cell.get_col()][target_cell.get_row() - 1] = min
        elif func == 'COUNTIF':
            count = self.calculate_count_if(indexes, target_cell, criteria)
            target_cell.change_value(str(count))
            self.value_cols[target_cell.get_col()][target_cell.get_row() - 1] = count
        elif func == 'SQUARE':
            square = self.calculate_square(indexes[0], target_cell)
            target_cell.change_value(str(square))
            self.value_cols[target_cell.get_col()][target_cell.get_row() - 1] = square
        elif func == 'SQRT':
            sqrt = self.calculate_sqrt(indexes[0], target_cell)
            target_cell.change_value(str(sqrt))
            self.value_cols[target_cell.get_col()][target_cell.get_row() - 1] = sqrt
        else:
            print("The function is not valid.")
        self.data = pd.DataFrame(self.value_cols)
