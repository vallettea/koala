
import re
from collections import OrderedDict
from koala.excel.excelutil import col2num, num2col, flatten, is_number
import string

CELL_REF_RE = re.compile(r"\!?(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

def find_associated_values(ref, first = None, second = None):
    valid = False

    try:
        found = re.search(CELL_REF_RE, ref)
        col = found.group(1)
        row = found.group(2)

    except:
        raise Exception('Couldn\'t find match in cell ref')
    
    if type(first) == Range:
        for key, value in first.items():
            r, c = key
            if r == row or c == col:
                first_value = value
                valid = True
                break

        if not valid:
            raise Exception('First argument of Range operation is not valid')
    else:
        first_value = first

    valid = False

    if type(second) == Range:
        for key, value in second.items():
            r, c = key
            if r == row or c == col:
                second_value = value
                valid = True
                break

        if not valid:
            raise Exception('Second argument of Range operation is not valid')
    else:
        second_value = second
    
    return (first_value, second_value)

def check_value(a):
    try: # This is to avoid None or Exception returned by Range operations
        if float(a):
            return a
        else:
            return 0
    except:
        return 0

class Range(OrderedDict):

    def __init__(self, cells, values):
        if len(cells) != len(values):
            raise ValueError("cells and values in a Range must have the same size")

        result = []
        cleaned_cells = []

        cells = list(flatten(cells))
        values = list(flatten(values))

        for index, cell in enumerate(cells):
            found = re.search(CELL_REF_RE, cell)
            col = found.group(1)
            row = found.group(2)

            if '!' in cell:
                cleaned_cell = cell.split('!')[1]
            else:
                cleaned_cell = cell

            cleaned_cells.append(cleaned_cell)
            try: # Might not be needed
                result.append(((row, col), values[index]))
            except:
                result.append(((row, col), None))

        # cells ref need to be cleaned of sheet name => WARNING, sheet ref is lost !!!
        cells = cleaned_cells
        self.cells = cells # this is used to be able to reconstruct Ranges from results of Range operations
        self.length = len(cells)
        
        # get last cell
        last = cells[self.length - 1]
        first = cells[0]

        last_found = re.search(CELL_REF_RE, last)
        first_found = re.search(CELL_REF_RE, first)

        last_col = last_found.group(1)
        last_row = last_found.group(2)

        first_col = first_found.group(1)
        first_row = first_found.group(2)

        self.nb_cols = int(col2num(last_col)) - int(col2num(first_col)) + 1
        self.nb_rows = int(last_row) - int(first_row) + 1

        OrderedDict.__init__(self, result)

    def is_associated(self, other):
        if self.length != other.length:
            return None

        nb_v = 0
        nb_c = 0

        for index, key in enumerate(self.keys()):
            r1, c1 = key
            r2, c2 = other.keys()[index]

            if r1 == r2:
                nb_v += 1
            if c1 == c2:
                nb_c += 1

        if nb_v == self.length:
            return 'v'
        elif nb_c == self.length:
            return 'c'
        else:
            return None

    def get(self, row, col = None):
        nr = self.nb_rows
        nc = self.nb_cols

        values = self.values()
        cells = self.cells

        if nr == 1 or nc == 1: # 1-dim range
            if col is not None:
                raise ValueError('Trying to access 1-dim range value with 2 coordinates')
            else:
                return self.values()[row - 1]
            
        else: # could be optimised
            indices = range(len(values))

            if row == 0: # get column
                filtered_indices = filter(lambda x: x % nc == col - 1, indices)

                filtered_values = map(lambda i: values[i], filtered_indices)
                filtered_cells = map(lambda i: cells[i], filtered_indices)

                return Range(filtered_cells, filtered_values)

            elif col == 0: # get row

                filtered_indices = filter(lambda x: (x / nc) == row - 1, indices)

                filtered_values = map(lambda i: values[i], filtered_indices)
                filtered_cells = map(lambda i: cells[i], filtered_indices)

                return Range(filtered_cells, filtered_values)

            else:
                base_col_number = col2num(cells[0][0])
                new_ref = num2col(col + base_col_number - 1) + str(row)
                new_value = values[(row - 1)* nc + (col - 1)]

                return new_value
                # return Range([new_ref], [new_value])

    @staticmethod
    def apply_one(func, self, other, ref):
        function = func_dict[func]

        first, second = find_associated_values(ref, self, other)

        return function(first, second)

    @staticmethod
    def apply_all(func, self, other, ref = None):
        function = func_dict[func]

        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): function(value, other.values()[key]), enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): function(value, other), enumerate(self.values())))


    @staticmethod
    def add(a, b):
        try:
            return check_value(a) + check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def substract(a, b):
        try:
            return check_value(a) - check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def minus(a, b = None):
        # b is not used, but needed in the signature. Maybe could be better
        try:
            return -check_value(a)
        except Exception as e:
            return e

    @staticmethod
    def multiply(a, b):
        try:
            return check_value(a) * check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def divide(a, b):
        try:
            return check_value(a) / check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def is_equal(a, b):
        try:
            return check_value(a) == check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def is_not_equal(a, b):
        try:
            return check_value(a) != check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def is_strictly_superior(a, b):
        try:
            return check_value(a) > check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def is_strictly_inferior(a, b):
        try:
            return check_value(a) < check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def is_superior_or_equal(a, b):
        try:
            return check_value(a) >= check_value(b)
        except Exception as e:
            return e

    @staticmethod
    def is_inferior_or_equal(a, b):
        try:
            return check_value(a) <= check_value(b)
        except Exception as e:
            return e


func_dict = {
    "multiply": Range.multiply,
    "divide": Range.divide,
    "add": Range.add,
    "substract": Range.substract,
    "minus": Range.minus,
    "is_equal": Range.is_equal,
    "is_not_equal": Range.is_not_equal,
    "is_strictly_superior": Range.is_strictly_superior,
    "is_strictly_inferior": Range.is_strictly_inferior,
    "is_superior_or_equal": Range.is_superior_or_equal,
    "is_inferior_or_equal": Range.is_inferior_or_equal,
}