import re

def get_column_letters(n, zero_indexed = True):
    if zero_indexed: 
        n += 1

    string = ""
    
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def letter_to_index(letter):
    return ord(letter.lower()) - ord('a') + 1

def get_column_index(s, go = False):
    if not go:
        return get_column_index(s, go=True) - 1
    else:
        if len(s) == 0:
            return 0

        return letter_to_index(s[-1]) + 26*get_column_index(s[:-1], go=True)

def get_cell_index(*args):
    if type(args[0]) == tuple:
        get_cell_index(*args[0])
    row, column = args[0], args[1]
    return get_column_letters(column, zero_indexed = True) + str(row + 1)

def get_cell_coordinates(cell_index):
    column_letter = re.search(r'[a-zA-Z]+', cell_index)[0]

    column = get_column_index(column_letter)
    row = int(re.search(r'[0-9]+', cell_index)[0]) - 1

    return (row, column)

def closed_range(a, b):
    start = None
    stop = None
    
    if a < b:
        start = a
        stop = b
    else:
        start = b
        stop = a

    return range(start, stop + 1)


def convert_coordinates_from_negative(rows, columns, row, column):
    if row < 0:
        row = rows + row
    if column < 0:
        column = columns + column
    return row, column

def normalize_cell_notation(rows, columns, cell, column=None):
    row = cell

    if type(cell) == tuple:
        row, column = cell
    elif type(cell) == str:
        row, column = get_cell_coordinates(cell)

    row, column = convert_coordinates_from_negative(rows, columns, row, column)

    return row, column


def compute_formula(self, value):
    try:
        value = self.__process_formula(value)
    except TypeError:
        value = self.__process_formula(value, number_based = False)
    except SyntaxError:
        # Tell the user their syntax was off
        pass

    return value