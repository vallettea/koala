
import re
from collections import OrderedDict

CELL_REF_RE = re.compile(r"\!?(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

def get_values(ref, first, second = None):
    first_value = None
    second_value = second

    try:
        col = re.search(CELL_REF_RE, ref).group(1)
        row = re.search(CELL_REF_RE, ref).group(2)
    except:
        raise Exception('Couldn\'t find match in cell ref')
    
    for key, value in first.items():
        r, c = key
        if r == row or c == col:
            first_value = value
            break

    if first_value is None:
        raise Exception('First argument of Range operation is not valid')

    if type(second) == Range:
        for key, value in second.items():
            r, c = key
            if r == row or c == col:
                second_value = value
                break

        if second_value is None:
            raise Exception('Second argument of Range operation is not valid')
    
    return (first_value, second_value)

class Range(OrderedDict):

    def __init__(self, cells, values):
        result = []

        for index, cell in enumerate(cells):
            col = re.search(CELL_REF_RE, cell).group(1)
            row = re.search(CELL_REF_RE, cell).group(2)
            result.append(((row, col), values[index]))

        self.cells = cells # this is used to be able to reconstruct Ranges from results of Range operations

        OrderedDict.__init__(self, result)

    # CAUTION, for now, only 1 dimension ranges are supported

    def add(self, other, ref):
        first, second = get_values(ref, self, other)

        return first + second

    def substract(self, other, ref):
        first, second = get_values(ref, self, other)

        return first - second

    def multiply_one(self, other, ref):
        first, second = get_values(ref, self, other)

        return first * second

    def multiply_all(self, other, ref = None):
        return Range(self.cells, map(lambda (key, value): value * other.values()[key], enumerate(self.values())))

    def divide(self, other, ref):
        first, second = get_values(ref, self, other)

        return first / second

    def is_equal(self, other, ref):
        first, second = get_values(ref, self, other)

        return first == second

    def is_not_equal(self, other, ref):
        first, second = get_values(ref, self, other)

        return first != second

    def is_strictly_superior(self, other, ref):
        first, second = get_values(ref, self, other)

        return first > second

    def is_strictly_superior_all(self, other, ref):
        return Range(self.keys(), map(lambda (key, value): value > other.values()[key], enumerate(self.values())))

    def is_strictly_inferior(self, other, ref):
        first, second = get_values(ref, self, other)

        return first < second

    def is_strictly_inferior_all(self, other, ref):
        if type(other) == Range:
            return Range(self.cells, map(lambda (key, value): value < other.values()[key], enumerate(self.values())))
        else:
            return Range(self.cells, map(lambda (key, value): value < other, enumerate(self.values())))

    def is_superior_or_equal(self, other, ref):
        first, second = get_values(ref, self, other)

        return first >= second

    def is_inferior_or_equal(self, other, ref):
        first, second = get_values(ref, self, other)

        return first <= second
    
