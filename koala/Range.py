import re
import string

from ExcelError import ExcelError
from utils import *
from Cell import Cell

# WARNING: Range should never be imported directly. Import Range from excelutils instead.

### Range Utils ###

cache = {}

def parse_cell_address(ref):
    try:
        if ref not in cache:
            found = re.search(CELL_REF_RE, ref)
            col = found.group(1)
            row = found.group(2)
            result = (int(row), col)
            cache[ref] = result
            return result
        else:
            return cache[ref]
    except:
        raise Exception('Couldn\'t find match in cell ref')
    
def get_cell_address(sheet, tuple):
    row = tuple[0]
    col = tuple[1]

    if sheet is not None:
        return sheet + '!' + col + str(row)
    else:
        return col + str(row)

def check_value(a):
    try: # This is to avoid None or Exception returned by Range operations
        if float(a) or isinstance(a, (unicode, str)):
            return a
        else:
            return 0
    except:
        if a == 'True':
            return True
        elif a == 'False':
            return False
        else:
            return 0

class RangeCore(dict):

    def __init__(self, reference, values = None, cellmap = None, nrows = None, ncols = None, name = None):
        
        if type(reference) == list: # some Range calculations such as excellib.countifs() use filtered keys
            cells = reference
        else:
            reference = reference.replace('$','')
            try:
                cells, nrows, ncols = resolve_range(reference)
            except:
                return ValueError('Range ERROR') # Will still be considered as a Range object, since we are inside __init__...

        cells = list(flatten(cells))

        origin = parse_cell_address(cells[0]) if len(cells) > 0 else None # origin of Range

        if cellmap:
            cells = [cell for cell in cells if cell in cellmap]

        if values:
            if len(cells) != len(values):
                raise ValueError("cells and values in a Range must have the same size")

        try:
            sheet = cells[0].split('!')[0]
        except:
            sheet = None

        result = []
        order = []

        for index, cell in enumerate(cells):
            row,col = parse_cell_address(cell)
            order.append((row, col))
            try:
                if cellmap:
                    result.append(((row, col), cellmap[cell]))

                else:
                    if isinstance(values[index], RangeCore):
                        raise Exception('Range can\'t be values of Range', reference)
                    result.append(((row, col), values[index]))

            except: # when you don't provide any values
                result.append(((row, col), None))

        # dont allow messing with these params
        self.__cellmap = cellmap
        self.__name = reference if type(reference) != list else name
        self.__origin = origin
        self.__addresses = cells
        self.__order = order
        self.__length = len(cells)
        self.__nrows = nrows
        self.__ncols = ncols

        if ncols == 1 and nrows == 1:
            self.__type = 'scalar'
        elif ncols == 1:
            self.__type = 'vertical'
        elif nrows == 1:
            self.__type = 'horizontal'
        else:
            self.__type = 'bidimensional'
        self.__sheet = sheet

        dict.__init__(self, result)

    @property
    def name(self):
        return self.__name
    @property
    def origin(self):
        return self.__origin
    @property
    def addresses(self):
        return self.__addresses
    @property
    def order(self):
        return self.__order
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
    def values(self):
        if self.__cellmap:
            values = []
            for cell in self.cells:
                values.append(cell.value)
            return values
        else:
            return self.cells
    
    @values.setter
    def values(self, new_values):
        if self.__cellmap:
            for index, cell in enumerate(self.cells):
                if index < len(new_values):
                    cell.value = new_values[index]
        else:
            for index, value in enumerate(self.order):
                if index < len(new_values):
                    self[value] = new_values[index]

    @property
    def cells(self):
        return map(lambda c: self[c], self.order)

    def get(self, row, col = None):
        nr = self.nrows
        nc = self.ncols

        values = self.values
        cells = self.addresses

        if nr == 1 or nc == 1: # 1-dim range
            if col is not None:
                raise Exception('Trying to access 1-dim range value with 2 coordinates')
            else:
                return values[row - 1]
            
        else: # could be optimised
            origin_col = col2num(self.origin[1])
            origin_row = self.origin[0]

            if row == 0: # get column

                out_col = num2col(int(col2num(self.origin[1]) + col - 1))

                tuples = [(r, out_col) for r in range(origin_row, origin_row + self.nrows)]

                cells = []
                values = []

                for t in tuples:
                    if t in self:
                        values.append(self[t].value)
                    else:
                        values.append(None)
                    cells.append(get_cell_address(self.sheet, t))

                return RangeCore(cells, values = values, nrows = len(cells), ncols = 1)

            elif col == 0: # get row

                out_row = self.origin[0] + row - 1

                tuples = [(out_row, c) for c in range(origin_col, origin_col + self.ncols)]

                cells = []
                values = []

                for t in tuples:
                    if t in self:
                        values.append(self[t].value)
                    else:
                        values.append(None)
                    cells.append(get_cell_address(self.sheet, t))

                return RangeCore(cells, values = values, nrows = 1, ncols = len(cells))

            else:
                base_col_number = col2num(cells[0][0])
                new_ref = num2col(col + base_col_number - 1) + str(row)
                new_value = values[(row - 1)* nc + (col - 1)]

                return new_value

    @staticmethod
    def filter(range, bool_range):

        if range.type == 'bidimensional':
            raise Exception('Cant use filter on bidimensional Ranges')

        filtered_addresses = []
        filtered_values = []

        for index, value in enumerate(range.values):
            test_value = bool_range.values[index]

            if type(test_value) != bool:
                raise Exception('RangeCore.filter must be used with bool Range as a second argument')

            if test_value:
                filtered_addresses.append(range.addresses[index])
                filtered_values.append(value)

        ncols = 1
        nrows = 1

        if range.type == 'vertical':
            nrows = len(filtered_values)
        elif range.type == 'horizontal':
            ncols = len(filtered_values)

        return RangeCore(filtered_addresses, filtered_values, nrows = nrows, ncols = ncols)

    @staticmethod
    def find_associated_cell(ref, range):
        # This function retrieves the cell associated to ref in a Range
        # For instance, in the range [A1, B1, C1], the cell associated to B2 is B1
        # This is useful to mimic the way Excel works

        if ref is not None:
            row, col = ref

            if (range.length) == 0: # if a Range is empty, it means normally that all its cells are empty
                return None
            elif range.type == "vertical":
                if (row, range.origin[1]) in range.order:
                    return range.addresses[range.order.index((row, range.origin[1]))]
                else:
                    return None
            elif range.type == "horizontal":
                if (range.origin[0], col) in range.order:
                    return range.addresses[range.order.index((range.origin[0], col))]
                else:
                    return None
            elif range.type == "scalar":
                if (row, range.origin[1]) in range.order:
                    return range.addresses[range.order.index((row, range.origin[1]))]
                elif (range.origin[0], col) in range.order:
                    return range.addresses[range.order.index((range.origin[0], col))]
                elif (row, col) in range.order:
                    return range.addresses[range.order.index((row, col))]
                else:
                    return None
                
            else:
                return None
        else:
            return None
            
    @staticmethod
    def find_associated_value(ref, item):
        # This function is ALMOST equivalent to RangeCore.find_associated_cell, but retrieves the value and not the Cell.
        row, col = ref

        if isinstance(item, RangeCore):
            try:
                if (item.length) == 0: # if a Range is empty, it means normally that all its cells are empty
                    item_value = 0
                elif item.type == "vertical":
                    if item.__cellmap is not None:
                        item_value = item[(row, item.origin[1])].value if (row, item.origin[1]) in item else None
                    else:
                        item_value = item[(row, item.origin[1])] if (row, item.origin[1]) in item else None
                elif item.type == "horizontal":
                    
                    if item.__cellmap is not None:
                        try:
                            item_value = item[(item.origin[0], col)].value if (item.origin[0], col) in item else None
                        except:
                            raise Exception
                    else:
                        item_value = item[(item.origin[0], col)] if (item.origin[0], col) in item else None
                else:
                    raise ExcelError('#VALUE!', 'cannot use find_associated_value on %s' % item.type)
            except ExcelError as e:
                raise Exception('First argument of Range operation is not valid: ' + e.value)
        elif item is None:
            item_value = 0
        else:
            item_value = item

        return item_value

    @staticmethod
    def apply(func, first, second, ref = None):
        # This function decides whether RangeCore.apply_one or RangeCore.apply_all should be used
        # This is a necessary complement to what is decided in graph.py:OperandNode.emit()

        isAssociated = False

        if ref:
            if isinstance(first, RangeCore):
                if first.length == 0:
                    first = 0
                else:
                    isAssociated = RangeCore.find_associated_cell(ref, first) is not None
            elif second is not None and isinstance(second, RangeCore):
                if second.length == 0:
                    second = 0
                else:
                    isAssociated = RangeCore.find_associated_cell(ref, second) is not None
            else:
                isAssociated = False

            if isAssociated:
                return RangeCore.apply_one(func, first, second, ref)
            else:
                return RangeCore.apply_all(func, first, second, ref)
        else:
            return RangeCore.apply_all(func, first, second, ref)

    @staticmethod
    def apply_one(func, first, second, ref = None):
        # This function applies a function to range operands, only for the cells associated to ref
        # Note that non-range operands are supported by RangeCore.find_associated_value()

        function = func_dict[func]

        if ref is None:
            first_value = first
            second_value = second
        else:
            first_value = RangeCore.find_associated_value(ref, first)
            second_value = RangeCore.find_associated_value(ref, second)

        return function(first_value, second_value)

    @staticmethod
    def apply_all(func, first, second, ref = None):
        # This function applies a function to range operands, for all the cells in the Ranges

        function = func_dict[func]

        # Here, the first arg of RangeCore() has little importance: TBC
        if isinstance(first, RangeCore) and isinstance(second, RangeCore):
            if first.length != second.length:
                raise ExcelError('#VALUE!', 'apply_all must have 2 Ranges of identical length')
            
            vals = [function(
                x.value if type(x) == Cell else x,
                y.value if type(y) == Cell else y
            ) for x,y in zip(first.cells, second.cells)]

            return RangeCore(first.addresses, vals, nrows = first.nrows, ncols = first.ncols)
        
        elif isinstance(first, RangeCore):
            vals = [function(
                x.value if type(x) == Cell else x,
                second
            ) for x in first.cells]

            return RangeCore(first.addresses, vals, nrows = first.nrows, ncols = first.ncols)

        elif isinstance(second, RangeCore):
            vals = [function(
                first,
                x.value if type(x) == Cell else x
            ) for x in second.cells]

            return RangeCore(second.addresses, vals, nrows = second.nrows, ncols = second.ncols)

        else:
            return function(first, second)

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
            return ExcelError('#DIV/0!', e)

    @staticmethod
    def is_equal(a, b):
        try:            
            if not isinstance(a, (str, unicode)):
                a = check_value(a)
            if not isinstance(b, (str, unicode)):
                b = check_value(b)

            return a == b
        except Exception as e:
            return ExcelError('#N/A', e)

    @staticmethod
    def is_not_equal(a, b):
        try:
            if not isinstance(a, (str, unicode)):
                a = check_value(a)
            if not isinstance(a, (str, unicode)):
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
    "multiply": RangeCore.multiply,
    "divide": RangeCore.divide,
    "add": RangeCore.add,
    "substract": RangeCore.substract,
    "minus": RangeCore.minus,
    "is_equal": RangeCore.is_equal,
    "is_not_equal": RangeCore.is_not_equal,
    "is_strictly_superior": RangeCore.is_strictly_superior,
    "is_strictly_inferior": RangeCore.is_strictly_inferior,
    "is_superior_or_equal": RangeCore.is_superior_or_equal,
    "is_inferior_or_equal": RangeCore.is_inferior_or_equal,
}


def RangeFactory(cellmap = None):

    class Range(RangeCore):

        def __init__(self, reference, values = None):
            super(Range, self).__init__(reference, values, cellmap = cellmap)       

    return Range