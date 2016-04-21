
import re
from collections import OrderedDict
# from excelutils import col2num
import string
def col2num(col):
    
    if not col:
        raise Exception("Column may not be empty")
    
    tot = 0
    for i,c in enumerate([c for c in col[::-1] if c != "$"]):
        if c == '$': continue
        tot += (ord(c)-64) * 26 ** i
    return tot

def num2col(num):
    
    if num < 1:
        raise Exception("Number must be larger than 0: %s" % num)
    
    s = ''
    q = num
    while q > 0:
        (q,r) = divmod(q,26)
        if r == 0:
            q = q - 1
            r = 26
        s = string.ascii_uppercase[r-1] + s
    return s

CELL_REF_RE = re.compile(r"\!?(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

def get_values(ref, first = None, second = None):
    first_value = None
    second_value = None

    try:
        col = re.search(CELL_REF_RE, ref).group(1)
        row = re.search(CELL_REF_RE, ref).group(2)

    except:
        raise Exception('Couldn\'t find match in cell ref')
    
    if type(first) == Range:
        for key, value in first.items():
            r, c = key
            if r == row or c == col:
                first_value = value
                break

        if first_value is None:
            raise Exception('First argument of Range operation is not valid')
    else:
        first_value = first

    if type(second) == Range:
        for key, value in second.items():
            r, c = key
            if r == row or c == col:
                second_value = value
                break

        if second_value is None:
            raise Exception('Second argument of Range operation is not valid')
    else:
        second_value = second
    
    return (first_value, second_value)

class Range(OrderedDict):

    def __init__(self, cells, values):
        if len(cells) != len(values):
            raise ValueError("cells and values in a Range must have the same size")

        result = []
        cleaned_cells = []

        for index, cell in enumerate(cells):
            col = re.search(CELL_REF_RE, cell).group(1)
            row = re.search(CELL_REF_RE, cell).group(2)

            cleaned_cells.append(cell.split('!')[1])
            result.append(((row, col), values[index]))

        # cells ref need to be cleaned of sheet name => WARNING, sheet ref is lost !!!
        cells = cleaned_cells
        self.cells = cells # this is used to be able to reconstruct Ranges from results of Range operations
        self.length = len(cells)
        
        self.nb_cols = int(col2num(cells[self.length - 1][0])) - int(col2num(cells[0][0])) + 1

        # get last cell
        last = cells[self.length - 1]
        first = cells[0]

        self.nb_rows = int(cells[self.length - 1][1]) - int(cells[0][1]) + 1

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

                return Range([new_ref], [new_value])


    @staticmethod
    def add_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first + second

    @staticmethod
    def add_all(self, other, ref = None):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value + other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value + other, enumerate(self.values())))

    @staticmethod
    def substract_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first - second

    @staticmethod
    def substract_all(self, other, ref = None):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value - other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value - other, enumerate(self.values())))

    @staticmethod
    def multiply_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first * second

    @staticmethod
    def multiply_all(self, other, ref = None):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value * other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value * other, enumerate(self.values())))

    @staticmethod
    def divide_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first / second

    @staticmethod
    def divide_all(self, other, ref = None):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value / other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value / other, enumerate(self.values())))

    @staticmethod
    def is_equal_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first == second

    @staticmethod
    def is_equal_all(self, other, ref = None):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value == other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value == other, enumerate(self.values())))

    @staticmethod
    def is_not_equal_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first != second

    @staticmethod
    def is_not_equal_all(self, other, ref = None):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value != other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value != other, enumerate(self.values())))

    @staticmethod
    def is_strictly_superior_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first > second

    @staticmethod
    def is_strictly_superior_all(self, other, ref = None):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value > other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value > other, enumerate(self.values())))

    @staticmethod
    def is_strictly_inferior_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first < second

    @staticmethod
    def is_strictly_inferior_all(self, other, ref):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value < other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value < other, enumerate(self.values())))

    @staticmethod
    def is_superior_or_equal_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first >= second

    @staticmethod
    def is_superior_or_equal_all(self, other, ref):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value >= other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value >= other, enumerate(self.values())))

    @staticmethod
    def is_inferior_or_equal_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first <= second

    @staticmethod
    def is_inferior_or_equal_all(self, other, ref):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value <= other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value <= other, enumerate(self.values())))
    
