import json
import typing
import cell
import dict_of_cells
from typing import List, Tuple, Dict, Any, Union
import matplotlib.pyplot as plt
import sys
import re
import sheet
import argparse


def main():
    parser = argparse.ArgumentParser(description='This module is the main '
                                                 'module of the program. It '
                                                 'contains the main '
                                                 'functions of the '
                                                 'program.\nYou can open a '
                                                 'new file or load an '
                                                 'existing file. You can '
                                                 'add, update, delete, '
                                                 'calculate and save to a '
                                                 'file.\nYou can also save '
                                                 'to excel, csv, html and '
                                                 'pdf files.\nThis is the '
                                                 'way the input should be:\n')
    parser.add_argument('--add', help='ADD\nCOLUMN: <Column Name> <Row '
                                      'Number> <Value> , <Row Number> '
                                      '<Value> , ...\nROW: <Row number> '
                                      '<Column Name> <Value> , <Column Name> '
                                      '<Value> , ... \nCELL: <Column Name> '
                                      '<Row Number> <Value>\n')
    parser.add_argument('--update', help='UPDATE:\nCOLUMN: <Column Name> '
                                         '<Row Number> <Value> , '
                                         '<Row Number> <Value> , ...\nROW: '
                                         '<Row number> <Column Name> <Value> '
                                         ', <Column Name> <Value> , '
                                         '...\nCELL: <Column Name> <Row '
                                         'Number> <Value>\n')
    parser.add_argument('--calculate', help='CALCULATE: <Target Index> = '
                                            '<Function> <Indexes> / <Target '
                                            'Index> = <Cell> <Operator (*, '
                                            '+, :, -)> <Number> / <Cell>\n')
    parser.add_argument('--save', help='SAVE:\nFILE: <File Name>\nto EXCEL: '
                                       '<File Name>\nto CSV: <File Name>\nto '
                                       'HTML: <File Name>\nto PDF: <File '
                                       'Name>\n')
    parser.add_argument('--exit', help='EXIT\n')
    arguments = parser.parse_args()



"""
This module is the main module of the program. It contains the main functions of the program.
You can open a new file or load an existing file. You can add, update, delete, calculate and save to a file.
You can also save to excel, csv, html and pdf files.
This is the way the input should be:
ADD:
    COLUMN: <Column Name> <Row Number> <Value> , <Row Number> <Value> , ... 
    ROW: <Row number> <Column Name> <Value> , <Column Name> <Value> , ... 
    CELL: <Column Name> <Row Number> <Value>
UPDATE:
    COLUMN: <Column Name> <Row Number> <Value> , <Row Number> <Value> , ...
    ROW: <Row number> <Column Name> <Value> , <Column Name> <Value> , ...
    CELL: <Column Name> <Row Number> <Value>
CALCULATE: <Target Index> = <Function> <Indexes> / <Target Index> = <Cell> <Operator (*, +, :, -)> <Number> / <Cell>
SAVE:
    FILE: <File Name>
    to EXCEL: <File Name>
    to CSV: <File Name>
    to HTML: <File Name>
    to PDF: <File Name>
EXIT
"""


def save_file(sheet: sheet.Sheet, file_name: str) -> None:
    """
    This function gets a sheet and a file name and saves the sheet to a file.
    :param sheet: The sheet to save.
    :param file_name: The name of the file to save the sheet to.
    """
    file_name += '.json'
    with open(file_name, 'w') as write_file:
        json.dump(sheet.to_dict(), write_file)


def load_file(file_name: str) -> sheet.Sheet:
    """
    This function gets a file name and returns a sheet from the file.
    :param file_name: The name of the file to load the sheet from.
    :return: The sheet from the file.
    """
    with open(file_name, 'r') as read_file:
        data = json.load(read_file)
        return sheet.Sheet.from_dict(data)


def save_to_excel(sheet: sheet.Sheet, file_name: str) -> None:
    """
    This function gets a sheet and a file name and saves the sheet to an excel file.
    :param sheet: The sheet to save.
    :param file_name: The name of the file to save the sheet to.
    """
    file_name += '.xlsx'
    df = sheet.get_data()
    df.to_excel(file_name, index=False)


def save_to_csv(sheet: sheet.Sheet, file_name: str) -> None:
    """
    This function gets a sheet and a file name and saves the sheet to a csv file.
    :param sheet: The sheet to save.
    :param file_name: The name of the file to save the sheet to.
    """
    file_name += '.csv'
    df = sheet.get_data()
    df.to_csv(file_name, index=False)


def save_to_html(sheet: sheet.Sheet, file_name: str) -> None:
    """
    This function gets a sheet and a file name and saves the sheet to an html file.
    :param sheet: The sheet to save.
    :param file_name: The name of the file to save the sheet to.
    """
    file_name += '.html'
    df = sheet.get_data()
    df.to_html(file_name, index=False)


def save_to_pdf(sheet: sheet.Sheet, file_name: str) -> None:
    """
    This function gets a sheet and a file name and saves the sheet to a pdf file.
    :param sheet: The sheet to save.
    :param file_name: The name of the file to save the sheet to.
    """
    file_name += '.pdf'
    df = sheet.get_data()
    plt.figure(figsize=(10, 10))
    plt.table(cellText=df.values, colLabels=df.columns, cellLoc='center',
              loc='center')
    plt.axis('off')
    plt.savefig(file_name, format='pdf', bbox_inches='tight')
    plt.close()


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


def split_string(s: str) -> Tuple[Union[str, typing.Any]]:
    """
    This function gets a string and returns a tuple of the letter and the number in the string.
    :param s: The string to split.
    :return: The tuple of the letter and the number in the string.
    """
    match = re.match(r"([a-z]+)([0-9]+)", s, re.I)
    if match:
        items = match.groups()
    return items


def expand_list(lst: List[Tuple[str, int]]) -> List[Tuple[str, int]]:
    """
    This function gets a list of indexes and returns an expanded list of indexes.
    :param lst: The list of indexes.
    :return: The expanded list of indexes.
    """
    # Sort the list
    lst.sort()
    # Initialize the expanded list
    expanded_lst = []
    # Iterate over the sorted list
    for i in range(len(lst)):
        # If this is the first occurrence of the letter or the next letter is different
        if i == 0 or lst[i][0] != lst[i-1][0]:
            # Add all numbers from this one to the next one with the same letter
            expanded_lst.extend([(lst[i][0], j) for j in range(lst[i][1], lst[i+1][1] if i+1 < len(lst) and lst[i][0] == lst[i+1][0] else lst[i][1]+1)])
        elif i+1 == len(lst) or lst[i][0] != lst[i+1][0]:
            expanded_lst.extend([(lst[i][0], j) for j in range(lst[i-1][1]+1, lst[i][1]+1)])
    return expanded_lst


def calculate(sheet: sheet.Sheet, indexes: str, function: str, target_index: Tuple[str, int], criteria=None) -> None:
    """
    This function gets a sheet, a list of indexes, a target index and a function and adds to the sheet the result of the functino.
    :param target_index: The index to add the result to.
    :param sheet: The sheet to calculate.
    :param indexes: The indexes to calculate.
    :param function: The function to calculate.
    :return: None
    """
    if len(sheet.get_cols()) == 0:
        print("The sheet is empty.")
        return None
    if len(target_index) != 2:
        print("Invalid index.")
        return None
    if function not in ['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNTIF', 'SQUARE', 'SQRT']:
        print("Invalid function.")
        return None
    if len(indexes) == 0:
        print("Invalid index.")
        return None
    mult = False
    if ',' in indexes:
        split_indexes = indexes.split(',')
        temp_indexes = []
        for index in split_indexes:
            index = index.strip()
            temp_indexes += index.split(':')
        mult = True
    elif ':' in indexes:
        if indexes.count(':') > 1:
            print("Invalid index.")
            return None
        temp_indexes = indexes.split(':')
        mult = True
    if mult:
        index_list = []
        for index in temp_indexes:
            try:
                split_index = split_string(index)
                split_index = (split_index[0].upper(), split_index[1])
                index_list.append((split_index[0], int(split_index[1])))
            except:
                print("Invalid index.")
                return None
        expend = expand_list(index_list)
        temp = set(expend + index_list)
        index_list = list(temp)
    else:
        index_list = indexes.split(' + ')
        index_list = [(index[0].upper(), int(index[1])) for index in index_list]
    target_index = (target_index[0], int(target_index[1]))
    sheet.calculate_cell(index_list, target_index, function, criteria)


def convert_to_str(name_input: List[str]) -> str:
    """
    This function gets a list of strings and returns a string.
    :param name_input: The list of strings.
    :return: The string.
    """
    name = ""
    for i in name_input:
        name += i + " "
    return name.strip()


OUTPUT_MASSAGE = ('**************************************************************************************************************************\n'
                  'ADD: \n COLUMN: <Column Name> <Row Number> <Value> , <Row Number> <Value> , ... \n ROW: <Row number> '
                  '<Column Name> <Value> , <Column Name> <Value> , ... \n CELL: <Column Name> <Row Number> <Value> \n'
                  'UPDATE: \n COLUMN: <Column Name> <Row Number> <Value> , <Row Number> <Value> , ... \n ROW: <Row number> '
                  '<Column Name> <Value> , <Column Name> <Value> , ... \n CELL: <Column Name> <Row Number> <Value> \n'
                  'CALCULATE: <Target Index> = <Function> <Indexes> / <Target Index> = <Cell> <Operator (*, +, :, -)> <Number> / <Cell> \n'
                  'SAVE: \n FILE: <File Name> \n to EXCEL: <File Name> \n to CSV: <File Name> \n to HTML: <File Name> \n to PDF: <File Name> \n'
                  'EXIT \n'
                  '****************************************************************************************************************************')


if __name__ == '__main__':
    """
    Load file: / New file: <File Name>
    Commands:
    Save file: / Save to excel: / Save to csv: / Save to html: / Save to pdf: <File Name>
    Add column: / Add row: / Add cell: <Column Name> <Row Number> <Value>
    Update column: / Update row: / Update cell: <Column Name> <Row Number> <Value>
    Calculate: <Target Index> = <Function> <Indexes> / <Target Index> = <Cell> <Operator (*, +, :, -)> <Number> / <Cell>
    """
    main()
    args = sys.argv
    if len(sys.argv) > 1:
        user_input = args[1:]
    else:
        user_input = input("Do you want to load a file or to open a new "
                           "file? please enter your request and the file's "
                           "name: ").split()
    while True:
        try:
            if user_input[0] == 'load' and user_input[1] == 'file':
                try:
                    name = convert_to_str(user_input[2:]) + '.json'
                    sheet1 = load_file(name)
                    print(sheet1)
                    break
                except:
                    print("There is no such file.")
                    user_input = input("Do you want to load a file or to open a new "
                                       "file? please enter your request and the file's "
                                       "name: ").split()
            elif user_input[0] == 'new' and user_input[1] == 'file':
                name = convert_to_str(user_input[2:])
                sheet1 = sheet.Sheet(name)
                break
            else:
                print("Invalid input.")
                user_input = input("Do you want to load a file or to open a new "
                                   "file? please enter your request and the file's "
                                   "name: ").split()
        except:
            print("Invalid input.")
            user_input = input("Do you want to load a file or to open a new "
                               "file? please enter your request and the file's "
                               "name: ").split()
    user_input = input(OUTPUT_MASSAGE + '\n' + "What do you want to do? ").split()
    while args[0] != 'exit' and user_input[0] != 'exit':
        try:
            if user_input[0] == 'save':
                try:
                    if user_input[1] == 'file':
                        name = convert_to_str(user_input[2:])
                        save_file(sheet1, name)
                        print("File saved.")
                    elif user_input[2] == 'excel':
                        name = convert_to_str(user_input[3:])
                        save_to_excel(sheet1, name)
                        print("File saved.")
                    elif user_input[2] == 'csv':
                        name = convert_to_str(user_input[3:])
                        save_to_csv(sheet1, name)
                        print("File saved.")
                    elif user_input[2] == 'html':
                        name = convert_to_str(user_input[3:])
                        save_to_html(sheet1, name)
                        print("File saved.")
                    elif user_input[2] == 'pdf':
                        name = convert_to_str(user_input[3:])
                        save_to_pdf(sheet1, name)
                        print("File saved.")
                    else:
                        print("File wasn't saved, you chose an invalid option.")
                except:
                   print("Invalid input.")
            elif user_input[0] == 'add':
                try:
                    if user_input[1] == 'column':
                        try:
                            temp_dict = dict_of_cells.Dictionary_of_Cells(user_input[2])
                            i = 4
                            while i < len(user_input):
                                row = user_input[i - 1]
                                temp_list = []
                                while (i) < len(user_input) and user_input[i] != ',':
                                    temp_list.append(user_input[i])
                                    i += 1
                                value = convert_to_str(temp_list)
                                temp_dict.add_cell(
                                    cell.Cell(user_input[2], str(row), value))
                                i += 2
                            sheet1.add_column(temp_dict)
                        except:
                            print("Column wasn't added. Invalid input, please try again")
                    elif user_input[1] == 'row':
                        try:
                            temp_dict = {}
                            row = user_input[2]
                            try:
                                row = int(row)
                            except:
                                print("Row wasn't added. Invalid input, please try again")
                            i = 4
                            while i < len(user_input):
                                col = user_input[i - 1]
                                temp_list = []
                                while i < len(user_input) and user_input[i] != ',':
                                    temp_list.append(user_input[i])
                                    i += 1
                                value = convert_to_str(temp_list)
                                temp_dict[col.upper()] = cell.Cell(col, str(row), value)
                                i += 2
                            sheet1.add_row(temp_dict)
                        except:
                            print("Row wasn't added. Invalid input, please try again")
                    elif user_input[1] == 'cell':
                        try:
                            value = convert_to_str(user_input[4:])
                            sheet1.add_cell(cell.Cell(user_input[2], user_input[3], value))
                        except:
                            print("Cell wasn't added. Invalid input, please try again")
                    else:
                        print("You entered an invalid option.")
                except:
                   print("Invalid input.")
            elif user_input[0] == 'update':
                try:
                    if user_input[1] == 'cell':
                        try:
                            value = convert_to_str(user_input[4:])
                            cur_cell = sheet1.find_cell((user_input[2].upper(), int(user_input[3])))
                            sheet1.update_cell(cur_cell, value)
                        except:
                            print("Cell wasn't updated. Invalid input, please try again")
                    elif user_input[1] == 'column':
                        try:
                            temp_dict = dict_of_cells.Dictionary_of_Cells(user_input[2])
                            i = 4
                            while i < len(user_input):
                                row = user_input[i - 1]
                                temp_list = []
                                while i < len(user_input) and user_input[i] != ',':
                                    temp_list.append(user_input[i])
                                    i += 1
                                value = convert_to_str(temp_list)
                                temp_dict.add_cell(cell.Cell(user_input[2], str(row), value))
                                i += 2
                            sheet1.update_column(temp_dict)
                        except:
                            print("Column wasn't updated. Invalid input, please try again")
                    elif user_input[1] == 'row':
                        try:
                            temp_dict: Dict[str, cell.Cell] = {}
                            row = user_input[2]
                            try:
                                num_row = int(row)
                            except:
                                print("Row wasn't updated. Invalid input for row, please try again")
                            i = 4
                            while i < len(user_input):
                                col = user_input[i - 1]
                                temp_list = []
                                while i < len(user_input) and user_input[i] != ',':
                                    temp_list.append(user_input[i])
                                    i += 1
                                value = convert_to_str(temp_list)
                                temp_dict[col.upper()] = cell.Cell(col, str(num_row), value)
                                i += 2
                            sheet1.update_row(temp_dict)
                        except:
                            print("Row wasn't updated. Invalid input, please try again")
                    else:
                        print("You entered an invalid option.")
                except:
                    print("Invalid input.")
            elif user_input[0] == 'calculate':
                try:
                    if user_input[4] in ['*', '+', ':', '-']:
                        target_location = (user_input[1][0].upper(), int(user_input[1][1]))
                        expression = convert_to_str(user_input[3:]).upper()
                        cells = find_cells(expression)
                        cell_values = {}
                        for c in cells:
                            location = (c[0].upper(), int(c[1]))
                            if not sheet1.check_if_in(location):
                                print("You've entered a non existing index.")
                                raise ValueError
                            else:
                                cur_cell = sheet1.find_cell(location)
                                cell_values[c] = cur_cell.get_value()
                        try:
                            num_expression = replace_cells_with_values(expression, cells, cell_values)
                            value = sheet.calculate_expression(num_expression)
                        except:
                            print("You've entered an invalid expression.")
                        if sheet1.check_if_in((target_location[0], target_location[1])):
                            new_cell = sheet1.find_cell((target_location[0], target_location[1]))
                            sheet1.update_cell(new_cell, str(value))
                        else:
                            sheet1.add_cell(cell.Cell(target_location[0], str(target_location[1]), str(value)))
                            new_cell = sheet1.find_cell((target_location[0], target_location[1]))
                        new_cell.set_formula("= " + expression)
                        for c1 in cells:
                            cur_cell = sheet1.find_cell((c1[0].upper(), int(c1[1])))
                            cur_cell.add_related_cell(new_cell)
                    elif user_input[3].upper() in ['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNTIF', 'SQUARE', 'SQRT']:
                        target_location = (user_input[1][0].upper(), int(user_input[1][1]))
                        if user_input[3].upper() == 'COUNTIF':
                            indexes = convert_to_str(user_input[4:len(user_input) - 1])
                            if len(indexes) > 1 and not ':' and ',' not in indexes:
                                raise Exception("That's an invalid input for indexes. ")
                            criteria = user_input[-1]
                            calculate(sheet1, indexes, user_input[3].upper(), target_location, criteria)
                        else:
                            indexes = convert_to_str(user_input[4:])
                            if len(indexes) > 1 and not ':' and ',' not in indexes:
                                raise Exception(
                                    "That's an invalid input for indexes. ")
                            calculate(sheet1, indexes, user_input[3].upper(), target_location, None)
                    else:
                        print("Invalid function, there is not such option.")
                except:
                    print("Couldn't calculate, Invalid input.")
            elif user_input[0] == 'delete':
                try:
                    if user_input[1] == 'column':
                        try:
                            sheet1.delete_column(user_input[2].upper())
                        except:
                            print("Column wasn't deleted. Invalid input, please try again")
                    elif user_input[1] == 'row':
                        try:
                            sheet1.delete_row(int(user_input[2]))
                        except:
                            print("Row wasn't deleted. Invalid input, please try again")
                    elif user_input[1] == 'cell':
                        try:
                            sheet1.delete_cell((user_input[2].upper(), int(user_input[3])))
                        except:
                            print("Cell wasn't deleted. Invalid input, please try again")
                    else:
                        print("You entered an invalid option.")
                except:
                    print("Invalid input.")
            else:
                print("Invalid option.")
            print(sheet1)
        except:
            print("Invalid input.")
        user_input = input(OUTPUT_MASSAGE + '\n' + "What do you want to do? ").split()
        if user_input[0] == 'exit' or args[0] == 'exit':
            break


