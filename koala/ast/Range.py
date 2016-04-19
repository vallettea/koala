
import re
from collections import OrderedDict

CELL_REF_RE = re.compile(r"\!?(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

def get_values(ref, first, second = None):
    first_value = None
    second_value = None

    try:
        col = re.search(CELL_REF_RE, ref).group(1)
        row = re.search(CELL_REF_RE, ref).group(2)
    except:
        raise Exception('Couldn\'t find match in cell ref')
    
    for key, value in first.items():
        r, c = key
        if r == row or c == col:
            first_value = value

    if not first_value:
        raise Exception('First argument of Range operation is not valid')

    if second is not None:
        for key, value in second.items():
            r, c = key
            if r == row or c == col:
                second_value = value

        if not second_value:
            raise Exception('Second argument of Range operation is not valid')
    
    return (first_value, second_value)

class Range(OrderedDict):

    def __init__(self, cells, values):
        result = []

        for index, cell in enumerate(cells):
            col = re.search(CELL_REF_RE, cell).group(1)
            row = re.search(CELL_REF_RE, cell).group(2)
            result.append(((row, col), values[index]))

        OrderedDict.__init__(self, result)

    # CAUTION, for now, only 1 dimension ranges are supported

    def add(self, other, ref):
        if type(other) == Range:
            first, second = get_values(ref, self, other)
        else:
            first = get_values(ref, self)[0]
            second = other

        return first + second

    # def substract(self, other, ref):
    #     first, second = get_values(ref, self, other)

    #     return first - second

        # if type(other) == Range:
        #     return check_array(self, index) + check_array(other, index)
        # else:
        #     return check_array(self, index) + other

    # def substract(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) - check_array(other, index)
    #     else:
    #         return check_array(self, index) - other

    # def multiply(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) * check_array(other, index)
    #     else:
    #         return check_array(self, index) * other

    # def divide(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) / check_array(other, index)
    #     else:
    #         return check_array(self, index) / other

    # # not sure if this is needed:

    # # def OR(self, other, index):
    # #     if type(other) == Range:
    # #         return check_array(self, index) or check_array(other, index)
    # #     else:
    # #         return check_array(self, index) or other

    # # def AND(self, other, index):
    # #     if type(other) == Range:
    # #         return check_array(self, index) and check_array(other, index)
    # #     else:
    # #         return check_array(self, index) and other

    # def is_equal(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) == check_array(other, index)
    #     else:
    #         return check_array(self, index) == other

    # def is_not_equal(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) != check_array(other, index)
    #     else:
    #         return check_array(self, index) != other

    # def is_strictly_superior(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) > check_array(other, index)
    #     else:
    #         return check_array(self, index) > other

    # def is_strictly_inferior(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) < check_array(other, index)
    #     else:
    #         return check_array(self, index) < other

    # def is_superior_or_equal(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) >= check_array(other, index)
    #     else:
    #         return check_array(self, index) >= other

    # def is_inferior_or_equal(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) <= check_array(other, index)
    #     else:
    #         return check_array(self, index) <= other
    
