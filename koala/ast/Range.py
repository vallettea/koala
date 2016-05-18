
import re
from collections import OrderedDict, Iterable
import string
from ExcelError import ExcelError
from excelutils import *

# WARNING: Range should never be imported directly. Import Range from excelutils instead.

### Range Utils ###

CELL_REF_RE = re.compile(r"\!?(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

def parse_cell_address(ref):
    try:
        found = re.search(CELL_REF_RE, ref)
        col = found.group(1)
        row = found.group(2)

        return (row, col)
    except:
        raise Exception('Couldn\'t find match in cell ref')
    
def get_cell_address(sheet, tuple):
    row = tuple[0]
    col = tuple[1]

    if sheet is not None:
        return sheet + '!' + col + row
    else:
        return col + row

def find_associated_values(ref, first = None, second = None):

    row, col = ref

    if type(first) == Range:
        try:
            if first.type == "vertical":
                first_value = first[(str(row), first.start[1])]
            elif first.type == "horizontal":
                first_value = first[(first.start[0], str(col))]
            else:
                raise ExcelError('#VALUE!', 'cannot use find_associated_values on bidimensional.')
        except:
            raise Exception('First argument of Range operation is not valid')
    else:
        first_value = first


    if type(second) == Range:
        try:
            if second.type == "vertical":
                second_value = second[(str(row), second.start[1])]
            elif second.type == "horizontal":
                second_value = second[(second.start[0], str(col))]
            else:
                raise ExcelError('#VALUE!', 'cannot use find_associated_values on bidimensional.')
        except:
            raise Exception('First argument of Range operation is not valid')
    else:
        second_value = second
    
    return (first_value, second_value)

def check_value(a):
    try: # This is to avoid None or Exception returned by Range operations
        if float(a):
            if type(a) == float:
                return round(a, 10)
            else:
                return a
        else:
            return 0
    except:
        return 0

### Range class ###

class Range(OrderedDict):

    def __init__(self, address, values = None):

        try:
            cells, nrows, ncols = resolve_range(address)
        except:
            raise ValueError('Range must not be a scalar')

        cells = list(flatten(cells))

        if cells[0] == cells[len(cells) - 1]:
            raise ValueError('Range must not be a scalar')

        result = []

        if len(cells) != len(values):
            raise ValueError("cells and values in a Range must have the same size")

        try:
            sheet = cells[0].split('!')[0]
        except:
            sheet = None

        for index, cell in enumerate(cells):
            found = re.search(CELL_REF_RE, cell)
            col = found.group(1)
            row = found.group(2)
            
            try:
                result.append(((row, col), values[index]))
            except: # when you don't provide any values
                result.append(((row, col), None))

        # dont allow messing with these params
        self.__address = address.replace('$','')
        self.__cells = cells
        self.__length = len(cells)
        self.__nrows = nrows
        self.__ncols = ncols
        if ncols == 1:
            self.__type = 'vertical'
        elif nrows == 1:
            self.__type = 'horizontal'
        else:
            self.__type = 'bidimensional'
        self.__sheet = sheet
        self.__start = parse_cell_address(cells[0])

        OrderedDict.__init__(self, result)       
        
    def __str__(self):
        return self.__address 

    @property
    def address(self):
        return self.__address
    @property
    def cells(self):
        return self.__cells
    @property
    def length(self):
        return self.__length
    @property
    def nrows(self):
        return self.__nrows
    @property
    def ncols(self):
        return self.__ncols
    @property
    def type(self):
        return self.__type
    @property
    def sheet(self):
        return self.__sheet
    @property
    def start(self):
        return self.__start
    @property
    def value(self):
        return self.values()
    
    @value.setter
    def value(self, new_values):
        for index, key in enumerate(self.keys()):
            self[key] = new_values[index]

    def reset(self):
        for key in self.keys():
            self[key] = None

    # def is_associated(self, other):
    #     if self.length != other.length:
    #         return None

    #     nb_v = 0
    #     nb_c = 0

    #     for index, key in enumerate(self.keys()):
    #         r1, c1 = key
    #         r2, c2 = other.keys()[index]

    #         if r1 == r2:
    #             nb_v += 1
    #         if c1 == c2:
    #             nb_c += 1

    #     if nb_v == self.length:
    #         return 'v'
    #     elif nb_c == self.length:
    #         return 'c'
    #     else:
    #         return None

    def get(self, row, col = None):
        nr = self.nrows
        nc = self.ncols

        values = self.values()
        cells = self.cells

        if nr == 1 or nc == 1: # 1-dim range
            if col is not None:
                raise Exception('Trying to access 1-dim range value with 2 coordinates')
            else:
                return self.values()[row - 1]
            
        else: # could be optimised
            indices = range(len(values))

            if row == 0: # get column
                filtered_indices = filter(lambda x: x % nc == col - 1, indices)

                filtered_values = map(lambda i: values[i], filtered_indices)
                filtered_cells = map(lambda i: cells[i], filtered_indices)

                new_address = str(filtered_cells[0]) + ':' + str(filtered_cells[len(filtered_cells)-1])

                return Range(new_address, filtered_values)

            elif col == 0: # get row

                filtered_indices = filter(lambda x: (x / nc) == row - 1, indices)

                filtered_values = map(lambda i: values[i], filtered_indices)
                filtered_cells = map(lambda i: cells[i], filtered_indices)

                new_address = str(filtered_cells[0]) + ':' + str(filtered_cells[len(filtered_cells)-1])

                return Range(new_address, filtered_values)

            else:
                base_col_number = col2num(cells[0][0])
                new_ref = num2col(col + base_col_number - 1) + str(row)
                new_value = values[(row - 1)* nc + (col - 1)]

                return new_value

    @staticmethod
    def apply_one(func, self, other, ref = None):
        function = func_dict[func]

        if ref is None:
            first = self
            second = other
        else:
            first, second = find_associated_values(ref, self, other)

        return function(first, second)

    @staticmethod
    def apply_all(func, self, other, ref = None):
        function = func_dict[func]

        # Here, the first arg of Range() has little importance: TBC

        if type(self) == Range and type(other) == Range:
            if self.length != other.length:
                raise ExcelError('#VALUE!', 'apply_all must have 2 Ranges of identical length')
            return Range(self.address, map(lambda (key, value): function(value, other.values()[key]), enumerate(self.values())))

        elif type(self) == Range:
            return Range(self.address, map(lambda (key, value): function(value, other), enumerate(self.values())))
        elif type(other) == Range:
            return Range(other.address, map(lambda (key, value): function(value, other), enumerate(other.values())))
        else:
            return function(self, other)


    @staticmethod
    def add(a, b):
        try:
            return check_value(a) + check_value(b)
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def substract(a, b):
        try:
            return check_value(a) - check_value(b)
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def minus(a, b = None):
        # b is not used, but needed in the signature. Maybe could be better
        try:
            return -check_value(a)
        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def multiply(a, b):
        try:
            return check_value(a) * check_value(b)
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def divide(a, b):
        try:
            return float(check_value(a)) / float(check_value(b))
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def is_equal(a, b):
        try:            
            if type(a) != str:
                a = check_value(a)
            if type(b) != str:
                b = check_value(b)
            # if a == 'David':
            #     print 'Check value', check_value(a)


            return a == b
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def is_not_equal(a, b):
        try:
            if type(a) != str:
                a = check_value(a)
            if type(b) != str:
                b = check_value(b)

            return a != b
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def is_strictly_superior(a, b):
        try:
            return check_value(a) > check_value(b)
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def is_strictly_inferior(a, b):
        try:
            return check_value(a) < check_value(b)
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def is_superior_or_equal(a, b):
        try:
            return check_value(a) >= check_value(b)
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def is_inferior_or_equal(a, b):
        try:
            return check_value(a) <= check_value(b)
        except Exception as e:
            return ExcelError('#N/A', e)


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