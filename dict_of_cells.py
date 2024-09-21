from typing import List
import cell


class Dictionary_of_Cells:

    """
    This class is a dictionary of cells.
    attributes:
    name: The name of the column.
    cells: A list of cells.
    methods:
    __init__: The constructor of the class.
    __str__: The string representation of the class.
    add_cell: Adds a cell to the list of cells.
    update_cell: Updates the value of a cell.
    get_name: Returns the name of the column.
    get_cells: Returns the list of cells.
    """

    def __init__(self, name: str):
        name = name.upper()
        while name not in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':  # Checking validity for the column.
            raise ValueError("That's an invalid input for a column, please enter again. ")
        self.name = name
        self.cells: list[cell.Cell] = []

    def __str__(self: 'Dictionary_of_Cells') -> str:
        return self.name + " : " + str(self.cells)

    def add_cell(self, cell: cell.Cell) -> None:
        self.cells.append(cell)

    def update_cell(self, value: str, row: int) -> None:
        self.cells[row].change_value(value)

    def get_name(self) -> str:
        return self.name

    def get_cells(self) -> List[cell.Cell]:
        return self.cells



