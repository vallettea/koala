import re
import string
from ExcelError import ExcelError
from excelutils import *

# WARNING: Range should never be imported directly. Import Range from excelutils instead.

### Range Utils ###

CELL_REF_RE = re.compile(r"\!?(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

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
        if float(a):
            if type(a) == float:
                return round(a, 10)
            else:
                return a
        else:
            return 0
    except:
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
                raise ValueError('Range must not be a scalar')

        cells = list(flatten(cells))

        if cellmap:
            cells = [cell for cell in cells if cell in cellmap]

        # if len(cells) > 0 and cells[0] == cells[len(cells) - 1]:
        #     print 'WARNING Range is a scalar', address, cells

        # Fill the Range with cellmap values 
        # if cellmap:
        #     cells = [cell for cell in cells if cell in cellmap]

        #     values = []

        #     for cell in cells:
        #         if cell in cellmap: # this is to avoid Sheet1!A5 and other empty cells due to A:A style named range
        #             try:
        #                 if isinstance(cellmap[cell].value, RangeCore):
        #                     raise Exception('Range can\'t be values of Range')
        #                 values.append(cellmap[cell].value)
        #             except: # if cellmap is not filled with actual Cells (for tests for instance)
        #                 if isinstance(cellmap[cell], RangeCore):
        #                     raise Exception('Range can\'t be values of Range')
        #                 values.append(cellmap[cell])

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
        self.__addresses = cells
        self.order = order
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
        self.__start = parse_cell_address(cells[0]) if len(cells) > 0 else None

        dict.__init__(self, result)

    @property
    def name(self):
        return self.__name
    @property
    def addresses(self):
        return self.__addresses
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
                cell.value = new_values[index]
        else:
            for key, value in enumerate(self.order):
                self[value] = new_values[key]

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
            indices = range(len(values))

            if row == 0: # get column
                filtered_indices = filter(lambda x: x % nc == col - 1, indices)

                filtered_values = map(lambda i: values[i], filtered_indices)
                filtered_cells = map(lambda i: cells[i], filtered_indices)

                new_address = str(filtered_cells[0]) + ':' + str(filtered_cells[len(filtered_cells)-1])

                return RangeCore(new_address, filtered_values)

            elif col == 0: # get row

                filtered_indices = filter(lambda x: (x / nc) == row - 1, indices)

                filtered_values = map(lambda i: values[i], filtered_indices)
                filtered_cells = map(lambda i: cells[i], filtered_indices)

                new_address = str(filtered_cells[0]) + ':' + str(filtered_cells[len(filtered_cells)-1])

                return RangeCore(new_address, filtered_values)

            else:
                base_col_number = col2num(cells[0][0])
                new_ref = num2col(col + base_col_number - 1) + str(row)
                new_value = values[(row - 1)* nc + (col - 1)]

                return new_value

    @staticmethod
    def find_associated_cell(ref, range):
        # This function retrieves the cell associated to ref in a Range
        # For instance, in the range [A1, B1, C1], the cell associated to B2 is B1
        # This is useful to mimic the way Excel works

        row, col = ref

        if (range.length) == 0: # if a Range is empty, it means normally that all its cells are empty
            return None
        elif range.type == "vertical":
            if (row, range.start[1]) in range.order:
                return range.addresses[range.order.index((row, range.start[1]))]
            else:
                return None
        elif range.type == "horizontal":
            if (range.start[0], col) in range.order:
                return range.addresses[range.order.index((range.start[0], col))]
            else:
                return None
        else:
            raise ExcelError('#VALUE!', 'cannot use find_associated_cell on %s' % range.type)

    @staticmethod
    def find_associated_values(ref, first = None, second = None):
        # This function is equivalent to RangeCore.find_associated_values, but for up to 2 ranges, and retrieves the value and not the Cell.

        # WARNING:
        # This function might be deprecated in future evolutions since find_associated_cell might be more elegant
        # combined with the RangeCore.apply() strategy

        row, col = ref

        if isinstance(first, RangeCore):
            try:
                if (first.length) == 0: # if a Range is empty, it means normally that all its cells are empty
                    first_value = 0
                elif first.type == "vertical":
                    if first.__cellmap is not None:
                        first_value = first[(row, first.start[1])].value
                    else:
                        first_value = first[(row, first.start[1])]
                elif first.type == "horizontal":
                    if first.__cellmap is not None:
                        try:
                            first_value = first[(first.start[0], col)].value
                        except:
                            print 'WHAT', zip(first.order, first.values), (first.start[0], col)
                            raise Exception
                    else:
                        first_value = first[(first.start[0], col)]
                else:
                    raise ExcelError('#VALUE!', 'cannot use find_associated_values on %s' % first.type)
            except ExcelError as e:
                raise Exception('First argument of Range operation is not valid: ' + e)
        else:
            first_value = first


        if isinstance(second, RangeCore):
            try:
                if (second.length) == 0: # if a Range is empty, it means normally that all its cells are empty
                    second_value = 0
                elif second.type == "vertical":
                    if second.__cellmap is not None:
                        second_value = second[(row, second.start[1])].value
                    else:
                        second_value = second[(row, second.start[1])]
                elif second.type == "horizontal":
                    if second.__cellmap is not None:
                        second_value = second[(second.start[0], col)].value
                    else:
                        second_value = second[(second.start[0], col)]
                else:
                    raise ExcelError('#VALUE!', 'cannot use find_associated_values on %s' % second.type)

            except ExcelError as e:
                raise Exception('Second argument of Range operation is not valid: ' + e)
        else:
            second_value = second
        
        return (first_value, second_value)

    @staticmethod
    def apply(func, self, other, ref = None):
        # This function decides whether RangeCore.apply_one or RangeCore.apply_all should be used
        # This is a necessary complement to what is decided in graph.py:OperandNode.emit()
        
        if ref:
            if isinstance(self, RangeCore) and self.length > 0:
                cell = RangeCore.find_associated_cell(ref, self)
            elif other is not None and isinstance(other, RangeCore) and other.length > 0:
                cell = RangeCore.find_associated_cell(ref, other)
            else:
                cell = None

            if cell is not None:
                return RangeCore.apply_one(func, self, other, ref)
            else:
                return RangeCore.apply_all(func, self, other, ref)
        else:
            return RangeCore.apply_all(func, self, other, ref)

    @staticmethod
    def apply_one(func, self, other, ref = None):
        # This function applies a function to range operands, only for the cells associated to ref
        # Note that non-range operands are supported by RangeCore.find_associated_values()

        function = func_dict[func]

        if ref is None:
            first = self
            second = other
        else:
            first, second = RangeCore.find_associated_values(ref, self, other)

        return function(first, second)

    @staticmethod
    def apply_all(func, self, other, ref = None):
        # This function applies a function to range operands, for all the cells in the Ranges

        function = func_dict[func]

        # Here, the first arg of RangeCore() has little importance: TBC
        if isinstance(self, RangeCore) and isinstance(other, RangeCore):
            if self.length != other.length:
                raise ExcelError('#VALUE!', 'apply_all must have 2 Ranges of identical length')
            
            vals = [function(
                x.value if type(x) == Cell else x,
                y.value if type(y) == Cell else y
            ) for x,y in zip(self.cells, other.cells)]

            return RangeCore(self.addresses, vals, nrows = self.nrows, ncols = self.ncols)
        
        elif isinstance(self, RangeCore):
            vals = [function(
                x.value if type(x) == Cell else x,
                other
            ) for x in self.cells]
            
            return RangeCore(self.addresses, vals, nrows = self.nrows, ncols = self.ncols)

        elif isinstance(other, RangeCore):
            vals = [function(
                self,
                x.value if type(x) == Cell else x
            ) for x in other.cells]

            return RangeCore(other.addresses, vals, nrows = other.nrows, ncols = other.ncols)

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