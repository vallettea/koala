
import re
from collections import OrderedDict

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

        for index, cell in enumerate(cells):
            col = re.search(CELL_REF_RE, cell).group(1)
            row = re.search(CELL_REF_RE, cell).group(2)
            result.append(((row, col), values[index]))

        self.cells = cells # this is used to be able to reconstruct Ranges from results of Range operations
        self.length = len(cells)

        OrderedDict.__init__(self, result)

    # CAUTION, for now, only 1 dimension ranges are supported

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
    
