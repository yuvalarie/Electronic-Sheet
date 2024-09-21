import typing
from typing import Union, Dict
import pandas as pd


def calculate_expression(expression: str) -> typing.Any:
    """
    This function gets a string of a mathematical expression and returns the result of the expression.
    :param expression: a string of a mathematical expression.
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
    if any(char not in allowed_operators for char in expression if not char.isdigit()):
        print("Error: Illegal operator used.")
        return None
    try:
        # Evaluate the expression
        result = eval(expression)
    except SyntaxError:
        print("Error: Invalid syntax.")
        return None
    return result


def check_type(input: str) -> Union[str, int, float, None]:
    value: Union[str, int, float, None]
    """
    This function gets a string snd converts it to a number if it's a number, and returns a string if it's not.
    :param input: A string of input from the user.
    :return: The input after converted to its type.
    """
    if pd.isna(input):
        value = None
    else:
        try:
            value = float(input)
            if '.' in input:
                value = round(value, 2)
            else:
                value = int(input)  # Converting to int
            if value == -0.0:
                value = 0.0
            elif value == -0:
                value = 0
        except ValueError:
            if input[0] == '=':  # Checking if the input is a formula
                value = calculate_expression((input[1:]).strip())
                if type(value) is float:
                    value = round(value, 2)
                if value is None:
                    value = input
            else:
                value = input
    return value


class Cell:

    """
    This class is a cell.
    attributes:
    col: The column of the cell.
    row: The row of the cell.
    value: The value of the cell.
    related_cells: A list of cells that are related to this cell.
    formula: The way the cell was calculated.
    methods:
    __init__: The constructor of the class.
    change_value: Changes the value of the cell.
    get_col: Returns the column of the cell.
    get_value: Returns the value of the cell.
    get_row: Returns the row of the cell.
    get_related_cells: Returns the list of related cells.
    add_related_cell: Adds a cell to the list of related cells.
    get_formula: Returns the formula of the cell.
    set_formula: Sets the formula of the cell.
    to_dict: Returns a dictionary of the cell.
    from_dict: Returns a cell from a dictionary.
    """

    def __init__(self, col: str, row: str, value=None) -> None:
        """
        This function creates an object of the type cell.
        :param col: A letter (or more) representing the column (can be a lower case).
        :param row: A number representing the row (must be an int).
        :param value: A string of input from the user to be used as the cell's value.
        """
        col = col.upper()
        while col not in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':  # Checking validity for the column.
            raise ValueError("That's an invalid input for a column, please enter again. ")
        self.col = col
        while True:  # Checking validity for the row.
            try:
                num_row = int(row)
                if num_row < 1:
                    raise ValueError
                break
            except ValueError:
                raise ValueError("That's an invalid input for a row, please enter again. ")
        self.row = num_row
        value = check_type(value)  # Converting the value into its type.
        self.value = value
        self.related_cells: list[Cell] = []  # A list of cells that are related to this cell.
        self.formula = ""  # The way the cell was calculated.

    def change_value(self, value: Union[str, int, float]) -> None:
        """
        This function gets a value and changes the value in the cell to the new value.
        :param value: A string of input from the user.
        :return: None
        """
        value = check_type(str(value))
        self.value = value

    def get_col(self) -> str:
        """
        This function returns the column of the cell.
        :return: The column of the cell.
        """
        return self.col

    def get_value(self) -> typing.Any:
        """
        This function returns the value of the cell.
        :return: the value of the cell.
        """
        return self.value

    def get_row(self) -> int:
        """
        This function returns the row of the cell.
        :return: the row of the cell.
        """
        return self.row

    def get_related_cells(self) -> list['Cell']:
        """
        This function returns the list of related cells.
        :return: the list of related cells.
        """
        return self.related_cells

    def add_related_cell(self, cell: 'Cell') -> None:
        """
        This function gets a cell and adds it to the list of related cells.
        :param cell: the cell to be added to the list of related cells.
        :return: none
        """
        if type(cell) is not Cell:
            raise ValueError("The input must be a cell.")
        if cell not in self.related_cells:
            self.related_cells.append(cell)

    def get_formula(self) -> Union[str, None]:
        """
        This function returns the formula of the cell.
        :return: the formula of the cell.
        """
        return self.formula

    def set_formula(self, formula: str) -> None:
        """
        This function gets a formula and sets the formula of the cell to the new formula.
        :param formula: the formula to be set to the cell.
        :return: None
        """
        self.formula = formula

    def to_dict(self) -> Dict[str, typing.Any]:
        """
        This function returns a dictionary of the cell.
        :return: A dictionary of the cell.
        """
        return {'column': self.col, 'row': self.row, 'value': self.value,
                'formula': self.formula, 'related_cells': [cell.to_dict() for cell in self.related_cells]}

    @classmethod
    def from_dict(cls: typing.Any, data: Dict[str, typing.Any]) -> 'Cell':
        """
        This function returns a cell from a dictionary.
        :param data: A dictionary of the cell.
        :return: The cell.
        """
        c = Cell(data['column'], data['row'], str(data['value']))
        c.set_formula(data['formula'])
        for cell in data['related_cells']:
            c.add_related_cell(Cell.from_dict(cell))
        return c

