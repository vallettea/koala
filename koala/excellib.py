# cython: profile=True

'''
Python equivalents of various excel functions
'''

# source: https://github.com/dgorissen/pycel/blob/master/src/pycel/excellib.py

from __future__ import absolute_import, division

import itertools
import numpy as np
import scipy.optimize
import datetime
from math import log, ceil
from decimal import Decimal, ROUND_UP, ROUND_HALF_UP
from calendar import monthrange
from dateutil.relativedelta import relativedelta

from openpyxl.compat import unicode

from koala.utils import *
from koala.Range import RangeCore as Range
from koala.ExcelError import *
from functools import reduce

######################################################################################
# A dictionary that maps excel function names onto python equivalents. You should
# only add an entry to this map if the python name is different to the excel name
# (which it may need to be to  prevent conflicts with existing python functions
# with that name, e.g., max).

# So if excel defines a function foobar(), all you have to do is add a function
# called foobar to this module.  You only need to add it to the function map,
# if you want to use a different name in the python code.

# Note: some functions (if, pi, atan2, and, or, array, ...) are already taken care of
# in the FunctionNode code, so adding them here will have no effect.
FUNCTION_MAP = {
    "ln":"xlog",
    "min":"xmin",
    "min":"xmin",
    "max":"xmax",
    "sum":"xsum",
    "gammaln":"lgamma",
    "round": "xround"
}

IND_FUN = [
    "SUM",
    "MIN",
    "IF",
    "TAN",
    "ATAN2",
    "PI",
    "ARRAY",
    "ARRAYROW",
    "AND",
    "OR",
    "ALL",
    "VALUE",
    "LOG",
    "MAX",
    "SUMPRODUCT",
    "IRR",
    "MIN",
    "SUM",
    "CHOOSE",
    "SUMIF",
    "AVERAGE",
    "RIGHT",
    "INDEX",
    "LOOKUP",
    "LINEST",
    "NPV",
    "MATCH",
    "MOD",
    "COUNT",
    "COUNTA",
    "COUNTIF",
    "COUNTIFS",
    "MATCH",
    "LOOKUP",
    "INDEX",
    "AVERAGE",
    "SUMIFS",
    "ROUND",
    "ROWS",
    "COLUMNS",
    "MID",
    "DATE",
    "YEARFRAC",
    "ISNA",
    "ISBLANK",
    "ISTEXT",
    "OFFSET",
    "SUMPRODUCT",
    "IFERROR",
    "XIRR",
    "VLOOKUP",
    "VDB",
    "SLN",
    "XNPV",
    "PMT",
    "ROUNDUP",
    "POWER",
    "SQRT",
    "TODAY",
    "YEAR",
    "MONTH",
    "EOMONTH",
]

CELL_CHARACTER_LIMIT = 32767
EXCEL_EPOCH = datetime.datetime.strptime("1900-01-01", '%Y-%m-%d').date()

######################################################################################
# List of excel equivalent functions
# TODO: needs unit testing


def value(text):
    # make the distinction for naca numbers
    if text.find('.') > 0:
        return float(text)
    elif text.endswith('%'):
        text = text.replace('%', '')
        return float(text) / 100
    else:
        return int(text)


def xlog(a):
    if isinstance(a,(list,tuple,np.ndarray)):
        return [log(x) for x in flatten(a)]
    else:
        #print a
        return log(a)


def xmax(*args): # Excel reference: https://support.office.com/en-us/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098
    # ignore non numeric cells and boolean cells
    values = extract_numeric_values(*args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(values) < 1:
        return 0
    else:
        return max(values)


def xmin(*args): # Excel reference: https://support.office.com/en-us/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152
    # ignore non numeric cells and boolean cells
    values = extract_numeric_values(*args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(values) < 1:
        return 0
    else:
        return min(values)


def xsum(*args): # Excel reference: https://support.office.com/en-us/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89
    # ignore non numeric cells and boolean cells

    values = extract_numeric_values(*args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(values) < 1:
        return 0
    else:
        return sum(values)


def choose(index_num, *values): # Excel reference: https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc

    index = int(index_num)

    if index <= 0 or index > 254:
        return ExcelError('#VALUE!', '%s must be between 1 and 254' % str(index_num))
    elif index > len(values):
        return ExcelError('#VALUE!', '%s must not be larger than the number of values: %s' % (str(index_num), len(values)))
    else:
        return values[index - 1]


def sumif(range, criteria, sum_range = None): # Excel reference: https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b

    # WARNING:
    # - wildcards not supported
    # - doesn't really follow 2nd remark about sum_range length

    if not isinstance(range, Range):
        return TypeError('%s must be a Range' % str(range))

    if isinstance(criteria, Range) and not isinstance(criteria , (str, bool)): # ugly...
        return 0

    indexes = find_corresponding_index(range.values, criteria)

    if sum_range:
        if not isinstance(sum_range, Range):
            return TypeError('%s must be a Range' % str(sum_range))

        def f(x):
            return sum_range.values[x] if x < sum_range.length else 0

        return sum(map(f, indexes))

    else:
        return sum([range.values[x] for x in indexes])


def sumifs(*args):
    # Excel reference: https://support.office.com/en-us/article/
    #   sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b

    nb_criteria = (len(args)-1) / 2

    args = list(args)

    # input checks
    if nb_criteria == 0:
        return TypeError('At least one criteria and criteria range should be provided.')
    if int(nb_criteria) != nb_criteria:
        return TypeError('Number of criteria an criteria ranges should be equal.')
    nb_criteria = int(nb_criteria)

    # separate arguments
    sum_range = args[0]
    criteria_ranges = args[1::2]
    criteria = args[2::2]
    index = list(range(0, len(sum_range)))

    for i in range(nb_criteria):

        criteria_range = criteria_ranges[i]
        criterion = str(criteria[i])

        index_tmp = find_corresponding_index(criteria_range.values, criterion)
        index = np.intersect1d(index, index_tmp)

    sum_select = [sum_range.values[i] for i in index]
    res = sum(sum_select)

    return res


def average(*args): # Excel reference: https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6
    # ignore non numeric cells and boolean cells
    values = extract_numeric_values(*args)

    return sum(values) / len(values)


def right(text,n):
    #TODO: hack to deal with naca section numbers
    if isinstance(text, unicode) or isinstance(text,str):
        return text[-n:]
    else:
        # TODO: get rid of the decimal
        return str(int(text))[-n:]


def index(my_range, row, col = None): # Excel reference: https://support.office.com/en-us/article/INDEX-function-a5dcf0dd-996d-40a4-a822-b56b061328bd

    for i in [my_range, row, col]:
        if isinstance(i, ExcelError) or i in ErrorCodes:
            return i

    row = int(row) if row is not None else row
    col = int(col) if col is not None else col

    if isinstance(my_range, Range):
        cells = my_range.addresses
        nr = my_range.nrows
        nc = my_range.ncols
    else:
        cells, nr, nc = my_range
        if nr > 1 or nc > 1:
            a = np.array(cells)
            cells = a.flatten().tolist()

    nr = int(nr)
    nc = int(nc)

    if type(cells) != list:
        return ExcelError('#VALUE!', '%s must be a list' % str(cells))

    if row is not None and not is_number(row):
        return ExcelError('#VALUE!', '%s must be a number' % str(row))

    if row == 0 and col == 0:
        return ExcelError('#VALUE!', 'No index asked for Range')

    if col is None and nr == 1 and row <= nc:
        # special case where index is matched on row, and the second row input can be used as a col
        col = row
        row = None

    if row is not None and row > nr:
        return ExcelError('#VALUE!', 'Index %i out of range' % row)

    if nr == 1:
        col = row if col is None else col
        return cells[int(col) - 1]

    if nc == 1:
        return cells[int(row) - 1]

    else: # could be optimised
        if col is None or row is None:
            return ExcelError('#VALUE!', 'Range is 2 dimensional, can not reach value with 1 arg as None')

        if not is_number(col):
            return ExcelError('#VALUE!', '%s must be a number' % str(col))

        if col > nc:
            return ExcelError('#VALUE!', 'Index %i out of range' % col)

        indices = list(range(len(cells)))

        if row == 0: # get column
            filtered_indices = [x for x in indices if x % nc == col - 1]
            filtered_cells = [cells[i] for i in filtered_indices]

            return filtered_cells

        elif col == 0: # get row
            filtered_indices = [x for x in indices if int(x / nc) == row - 1]
            filtered_cells = [cells[i] for i in filtered_indices]

            return filtered_cells

        else:
            return cells[(row - 1)* nc + (col - 1)]


def lookup(value, lookup_range, result_range = None): # Excel reference: https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb

    # TODO
    if not isinstance(value,(int,float)):
        return Exception("Non numeric lookups (%s) not supported" % value)

    # TODO: note, may return the last equal value

    # index of the last numeric value
    lastnum = -1
    for i,v in enumerate(lookup_range.values):
        if isinstance(v,(int,float)):
            if v > value:
                break
            else:
                lastnum = i

    output_range = result_range.values if result_range is not None else lookup_range.values

    if lastnum < 0:
        return ExcelError('#VALUE!', 'No numeric data found in the lookup range')
    else:
        if i == 0:
            return ExcelError('#VALUE!', 'All values in the lookup range are bigger than %s' % value)
        else:
            if i >= len(lookup_range)-1:
                # return the biggest number smaller than value
                return output_range[lastnum]
            else:
                return output_range[i-1]


# NEEDS TEST
def linest(*args, **kwargs): # Excel reference: https://support.office.com/en-us/article/LINEST-function-84d7d0d9-6e50-4101-977a-fa7abf772b6d

    Y = list(args[0].values())
    X = list(args[1].values())

    if len(args) == 3:
        const = args[2]
        if isinstance(const,str):
            const = (const.lower() == "true")
    else:
        const = True

    degree = kwargs.get('degree',1)

    # build the vandermonde matrix
    A = np.vander(X, degree+1)

    if not const:
        # force the intercept to zero
        A[:,-1] = np.zeros((1,len(X)))

    # perform the fit
    (coefs, residuals, rank, sing_vals) = np.linalg.lstsq(A, Y)

    return coefs


def npv(rate, *values): # Excel reference: https://support.office.com/en-us/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568
    cashflow = list(flatten_list(list(values)))

    if is_not_number_input(rate):
        return numeric_error(rate, 'rate')

    if is_not_number_input(cashflow):
        return numeric_error(cashflow, 'values')

    if isinstance(cashflow, Range):
        cashflow = cashflow.values

    return sum([float(x)*(1+rate)**-(i+1) for (i,x) in enumerate(cashflow)])


def rows(array):
    """
    Function to find the number of rows in an array.
    Excel reference: https://support.office.com/en-ie/article/rows-function-b592593e-3fc2-47f2-bec1-bda493811597

    :param array: the array of which the rows should be counted.
    :return: the number of rows.
    """

    if isinstance(array, (float, int)):
        rows = 1  # special case for A1:A1 type ranges which for some reason only return an int/float
    elif array is None:
        rows = 1  # some A1:A1 ranges return None (issue with ref cell)
    else:
        rows = len(array.values)

    return rows


def columns(array):
    """
    Function to find the number of columns in an array.
    Excel reference: https://support.office.com/en-us/article/columns-function-4e8e7b4e-e603-43e8-b177-956088fa48ca

    :param array: the array of which the columns should be counted.
    :return: the number of columns.
    """

    return rows(array)


def match(lookup_value, lookup_range, match_type=1): # Excel reference: https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a

    if not isinstance(lookup_range, Range):
        return ExcelError('#VALUE!', 'Lookup_range is not a Range')

    def type_convert(value):
        if type(value) == str:
            value = value.lower()
        elif type(value) == int:
            value = float(value)
        elif value is None:
            value = 0

        return value;
    def type_convert_float(value):
        if is_number(value):
            value = float(value)
        else:
            value = None

        return value

    lookup_value = type_convert(lookup_value)

    range_values = [x for x in lookup_range.values if x is not None] # filter None values to avoid asc/desc order errors
    range_length = len(range_values)

    if match_type == 1:
        # Verify ascending sort

        posMax = -1
        for i in range(range_length):
            current = type_convert(range_values[i])

            if i < range_length - 1:
                if current > type_convert(range_values[i + 1]):
                    return ExcelError('#VALUE!', 'for match_type 1, lookup_range must be sorted ascending')
            if current <= lookup_value:
                posMax = i
        if posMax == -1:
            return ExcelError('#VALUE!','no result in lookup_range for match_type 1')
        return posMax +1 #Excel starts at 1

    elif match_type == 0:
        # No string wildcard
        try:
            if is_number(lookup_value):
                lookup_value = float(lookup_value)
                output = [type_convert_float(x) for x in range_values].index(lookup_value) + 1
            else:
                output = [str(x).lower() for x in range_values].index(lookup_value) + 1
            return output
        except:
            return ExcelError('#VALUE!', '%s not found' % lookup_value)

    elif match_type == -1:
        # Verify descending sort
        posMin = -1
        for i in range((range_length)):
            current = type_convert(range_values[i])

            if i is not range_length-1 and current < type_convert(range_values[i+1]):
               return ExcelError('#VALUE!','for match_type -1, lookup_range must be sorted descending')
            if current >= lookup_value:
               posMin = i
        if posMin == -1:
            return ExcelError('#VALUE!', 'no result in lookup_range for match_type -1')
        return posMin +1 #Excel starts at 1


def mod(nb, q): # Excel Reference: https://support.office.com/en-us/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3
    if not isinstance(nb, int):
        return ExcelError('#VALUE!', '%s is not an integer' % str(nb))
    elif not isinstance(q, int):
        return ExcelError('#VALUE!', '%s is not an integer' % str(q))
    else:
        return nb % q


def eomonth(start_date, months): # Excel reference: https://support.office.com/en-us/article/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628
    if not is_number(start_date):
        return ExcelError('#VALUE!', 'start_date %s must be a number' % str(start_date))
    if start_date < 0:
        return ExcelError('#VALUE!', 'start_date %s must be positive' % str(start_date))

    if not is_number(months):
        return ExcelError('#VALUE!', 'months %s must be a number' % str(months))

    y1, m1, d1 = date_from_int(start_date)
    start_date_d = datetime.date(year=y1, month=m1, day=d1)
    end_date_d = start_date_d + relativedelta(months=int(months))
    y2 = end_date_d.year
    m2 = end_date_d.month
    d2 = monthrange(y2, m2)[1]
    res = int(int_from_date(datetime.date(y2, m2, d2)))

    return res


def year(serial_number): # Excel reference: https://support.office.com/en-us/article/year-function-c64f017a-1354-490d-981f-578e8ec8d3b9
    if not is_number(serial_number):
        return ExcelError('#VALUE!', 'start_date %s must be a number' % str(serial_number))
    if serial_number < 0:
        return ExcelError('#VALUE!', 'start_date %s must be positive' % str(serial_number))

    y1, m1, d1 = date_from_int(serial_number)

    return y1


def month(serial_number): # Excel reference: https://support.office.com/en-us/article/month-function-579a2881-199b-48b2-ab90-ddba0eba86e8
    if not is_number(serial_number):
        return ExcelError('#VALUE!', 'start_date %s must be a number' % str(serial_number))
    if serial_number < 0:
        return ExcelError('#VALUE!', 'start_date %s must be positive' % str(serial_number))

    y1, m1, d1 = date_from_int(serial_number)

    return m1


def count(*args): # Excel reference: https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c
    l = list(args)

    total = 0

    for arg in l:
        if isinstance(arg, Range):
            total += len([x for x in arg.values if is_number(x) and type(x) is not bool]) # count inside a list
        elif is_number(arg): # int() is used for text representation of numbers
            total += 1

    return total


def counta(range):
    if isinstance(range, ExcelError) or range in ErrorCodes:
        if range.value == '#NULL':
            return 0
        else:
            return range # return the Excel Error
            # raise Exception('ExcelError other than #NULL passed to excellib.counta()')
    else:
        return len([x for x in range.values if x != None])


def countif(range, criteria): # Excel reference: https://support.office.com/en-us/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34

    # WARNING:
    # - wildcards not supported
    # - support of strings with >, <, <=, =>, <> not provided

    valid = find_corresponding_index(range.values, criteria)

    return len(valid)


def countifs(*args): # Excel reference: https://support.office.com/en-us/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842

    arg_list = list(args)
    l = len(arg_list)

    if l % 2 != 0:
        return ExcelError('#VALUE!', 'excellib.countifs() must have a pair number of arguments, here %d' % l)


    if l >= 2:
        indexes = find_corresponding_index(args[0].values, args[1]) # find indexes that match first layer of countif

        remaining_ranges = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 0] # get only ranges
        remaining_criteria = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 1] # get only criteria

        # verif that all Ranges are associated COULDNT MAKE THIS WORK CORRECTLY BECAUSE OF RECURSION
        # association_type = None

        # temp = [args[0]] + remaining_ranges

        # for index, range in enumerate(temp): # THIS IS SHIT, but works ok
        #     if type(range) == Range and index < len(temp) - 1:
        #         asso_type = range.is_associated(temp[index + 1])

        #         print 'asso', asso_type
        #         if association_type is None:
        #             association_type = asso_type
        #         elif associated_type != asso_type:
        #             association_type = None
        #             break

        # print 'ASSO', association_type

        # if association_type is None:
        #     return ValueError('All items must be Ranges and associated')

        filtered_remaining_ranges = []

        for range in remaining_ranges: # filter items in remaining_ranges that match valid indexes from first countif layer
            filtered_remaining_cells = []
            filtered_remaining_range = []

            for index, item in enumerate(range.values):
                if index in indexes:
                    filtered_remaining_cells.append(range.addresses[index]) # reconstructing cells from indexes
                    filtered_remaining_range.append(item) # reconstructing values from indexes

            # WARNING HERE
            filtered_remaining_ranges.append(Range(filtered_remaining_cells, filtered_remaining_range))

        new_tuple = ()

        for index, range in enumerate(filtered_remaining_ranges): # rebuild the tuple that will be the argument of next layer
            new_tuple += (range, remaining_criteria[index])

        return min(countifs(*new_tuple), len(indexes)) # only consider the minimum number across all layer responses

    else:
        return float('inf')



def xround(number, num_digits = 0): # Excel reference: https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

    if not is_number(number):
        return ExcelError('#VALUE!', '%s is not a number' % str(number))
    if not is_number(num_digits):
        return ExcelError('#VALUE!', '%s is not a number' % str(num_digits))

    number = float(number) # if you don't Spreadsheet.dump/load, you might end up with Long numbers, which Decimal doesn't accept

    if num_digits >= 0: # round to the right side of the point
        return float(Decimal(repr(number)).quantize(Decimal(repr(pow(10, -num_digits))), rounding=ROUND_HALF_UP))
        # see https://docs.python.org/2/library/functions.html#round
        # and https://gist.github.com/ejamesc/cedc886c5f36e2d075c5

    else:
        return round(number, num_digits)


def roundup(number, num_digits = 0): # Excel reference: https://support.office.com/en-us/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7

    if not is_number(number):
        return ExcelError('#VALUE!', '%s is not a number' % str(number))
    if not is_number(num_digits):
        return ExcelError('#VALUE!', '%s is not a number' % str(num_digits))

    number = float(number) # if you don't Spreadsheet.dump/load, you might end up with Long numbers, which Decimal doesn't accept

    if num_digits >= 0: # round to the right side of the point
        return float(Decimal(repr(number)).quantize(Decimal(repr(pow(10, -num_digits))), rounding=ROUND_UP))
        # see https://docs.python.org/2/library/functions.html#round
        # and https://gist.github.com/ejamesc/cedc886c5f36e2d075c5

    else:
        return ceil(number / pow(10, -num_digits)) * pow(10, -num_digits)


def mid(text, start_num, num_chars): # Excel reference: https://support.office.com/en-us/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028

    text = str(text)

    if len(text) > CELL_CHARACTER_LIMIT:
        return ExcelError('#VALUE!', 'text is too long. Is %s needs to be %s or less.' % (len(text), CELL_CHARACTER_LIMIT))

    if type(start_num) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(start_num))
    if type(num_chars) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(num_chars))

    if start_num < 1:
        return ExcelError('#VALUE!', '%s is < 1' % str(start_num))
    if num_chars < 0:
        return ExcelError('#VALUE!', '%s is < 0' % str(num_chars))

    return text[(start_num - 1): (start_num - 1 + num_chars)]


def date(year, month, day): # Excel reference: https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349

    if type(year) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(year))

    if type(month) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(month))

    if type(day) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(day))

    if year < 0 or year > 9999:
        return ExcelError('#VALUE!', 'Year must be between 1 and 9999, instead %s' % str(year))

    if year < 1900:
        year = 1900 + year

    year, month, day = normalize_year(year, month, day) # taking into account negative month and day values

    date_0 = datetime.datetime(1900, 1, 1)
    date = datetime.datetime(year, month, day)

    result = (datetime.datetime(year, month, day) - date_0).days + 2

    if result <= 0:
        return ExcelError('#VALUE!', 'Date result is negative')
    else:
        return result


def yearfrac(start_date, end_date, basis = 0): # Excel reference: https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8

    def actual_nb_days_ISDA(start, end): # needed to separate days_in_leap_year from days_not_leap_year
        y1, m1, d1 = start
        y2, m2, d2 = end

        days_in_leap_year = 0
        days_not_in_leap_year = 0

        year_range = list(range(y1, y2 + 1))

        for y in year_range:

            if y == y1 and y == y2:
                nb_days = date(y2, m2, d2) - date(y1, m1, d1)
            elif y == y1:
                nb_days = date(y1 + 1, 1, 1) - date(y1, m1, d1)
            elif y == y2:
                nb_days = date(y2, m2, d2) - date(y2, 1, 1)
            else:
                nb_days = 366 if is_leap_year(y) else 365

            if is_leap_year(y):
                days_in_leap_year += nb_days
            else:
                days_not_in_leap_year += nb_days

        return (days_not_in_leap_year, days_in_leap_year)

    def actual_nb_days_AFB_alter(start, end): # http://svn.finmath.net/finmath%20lib/trunk/src/main/java/net/finmath/time/daycount/DayCountConvention_ACT_ACT_YEARFRAC.java
        y1, m1, d1 = start
        y2, m2, d2 = end

        delta = date(*end) - date(*start)

        if delta <= 365:
            if is_leap_year(y1) and is_leap_year(y2):
                denom = 366
            elif is_leap_year(y1) and date(y1, m1, d1) <= date(y1, 2, 29):
                denom = 366
            elif is_leap_year(y2) and date(y2, m2, d2) >= date(y2, 2, 29):
                denom = 366
            else:
                denom = 365
        else:
            year_range = list(range(y1, y2 + 1))
            nb = 0

            for y in year_range:
                nb += 366 if is_leap_year(y) else 365

            denom = nb / len(year_range)

        return delta / denom

    if not is_number(start_date):
        return ExcelError('#VALUE!', 'start_date %s must be a number' % str(start_date))
    if not is_number(end_date):
        return ExcelError('#VALUE!', 'end_date %s must be number' % str(end_date))
    if start_date < 0:
        return ExcelError('#VALUE!', 'start_date %s must be positive' % str(start_date))
    if end_date < 0:
        return ExcelError('#VALUE!', 'end_date %s must be positive' % str(end_date))

    if start_date > end_date: # switch dates if start_date > end_date
        temp = end_date
        end_date = start_date
        start_date = temp

    y1, m1, d1 = date_from_int(start_date)
    y2, m2, d2 = date_from_int(end_date)

    if basis == 0: # US 30/360
        d2 = 30 if d2 == 31 and (d1 == 31 or d1 == 30) else min(d2, 31)
        d1 = 30 if d1 == 31 else d1

        count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = count / 360

    elif basis == 1: # Actual/actual
        result = actual_nb_days_AFB_alter((y1, m1, d1), (y2, m2, d2))

    elif basis == 2: # Actual/360
        result = (end_date - start_date) / 360

    elif basis == 3: # Actual/365
        result = (end_date - start_date) / 365

    elif basis == 4: # Eurobond 30/360
        d2 = 30 if d2 == 31 else d2
        d1 = 30 if d1 == 31 else d1

        count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = count / 360

    else:
        return ExcelError('#VALUE!', '%d must be 0, 1, 2, 3 or 4' % basis)


    return result


def isna(value):
    # This function might need more solid testing
    try:
        eval(value)
        return False
    except:
        return True


def isblank(value):
    return value is None


def istext(value):
    return type(value) == str


def offset(reference, rows, cols, height=None, width=None): # Excel reference: https://support.office.com/en-us/article/OFFSET-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66
    # This function accepts a list of addresses
    # Maybe think of passing a Range as first argument

    for i in [reference, rows, cols, height, width]:
        if isinstance(i, ExcelError) or i in ErrorCodes:
            return i

    rows = int(rows)
    cols = int(cols)

    # get first cell address of reference
    if is_range(reference):
        ref = resolve_range(reference, should_flatten = True)[0][0]
    else:
        ref = reference
    ref_sheet = ''
    end_address = ''

    if '!' in ref:
        ref_sheet = ref.split('!')[0] + '!'
        ref_cell = ref.split('!')[1]
    else:
        ref_cell = ref

    found = re.search(CELL_REF_RE, ref)
    new_col = col2num(found.group(1)) + cols
    new_row = int(found.group(2)) + rows

    if new_row <= 0 or new_col <= 0:
        return ExcelError('#VALUE!', 'Offset is out of bounds')

    start_address = str(num2col(new_col)) + str(new_row)

    if (height is not None and width is not None):
        if type(height) != int:
            return ExcelError('#VALUE!', '%d must not be integer' % height)
        if type(width) != int:
            return ExcelError('#VALUE!', '%d must not be integer' % width)

        if height > 0:
            end_row = new_row + height - 1
        else:
            return ExcelError('#VALUE!', '%d must be strictly positive' % height)
        if width > 0:
            end_col = new_col + width - 1
        else:
            return ExcelError('#VALUE!', '%d must be strictly positive' % width)

        end_address = ':' + str(num2col(end_col)) + str(end_row)
    elif height and not width or not height and width:
        return ExcelError('Height and width must be passed together')

    return ref_sheet + start_address + end_address


def sumproduct(*ranges): # Excel reference: https://support.office.com/en-us/article/SUMPRODUCT-function-16753e75-9f68-4874-94ac-4d2145a2fd2e
    range_list = list(ranges)

    for r in range_list: # if a range has no values (i.e if it's empty)
        if len(r.values) == 0:
            return 0

    for range in range_list:
        for item in range.values:
            # If there is an ExcelError inside a Range, sumproduct should output an ExcelError
            if isinstance(item, ExcelError):
                return ExcelError("#N/A", "ExcelErrors are present in the sumproduct items")

    reduce(check_length, range_list) # check that all ranges have the same size

    return reduce(lambda X, Y: X + Y, reduce(lambda x, y: Range.apply_all('multiply', x, y), range_list).values)


def iferror(value, value_if_error): # Excel reference: https://support.office.com/en-us/article/IFERROR-function-c526fd07-caeb-47b8-8bb6-63f3e417f611

    if isinstance(value, ExcelError) or value in ErrorCodes:
        return value_if_error
    else:
        return value


def irr(values, guess = None):
    """
    Function to calculate the internal rate of return (IRR) using payments and periodic dates. It resembles the
    excel function IRR().

    Excel reference: https://support.office.com/en-us/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc

    :param values: the payments of which at least one has to be negative.
    :param guess: an initial guess which is required by Excel but isn't used by this function.
    :return: a float being the IRR.
    """
    if isinstance(values, Range):
        values = values.values

    if is_not_number_input(values):
        return numeric_error(values, 'values')

    if guess is not None and guess != 0:
        raise ValueError('guess value for excellib.irr() is %s and not 0' % guess)
    else:
        try:
            return np.irr(values)
        except Exception as e:
            return ExcelError('#NUM!', e)


def xirr(values, dates, guess=0):
    """
    Function to calculate the internal rate of return (IRR) using payments and non-periodic dates. It resembles the
    excel function XIRR().

    Excel reference: https://support.office.com/en-ie/article/xirr-function-de1242ec-6477-445b-b11b-a303ad9adc9d

    :param values: the payments of which at least one has to be negative.
    :param dates: the dates as excel dates (e.g. 43571 for 16/04/2019).
    :param guess: an initial guess which is required by Excel but isn't used by this function.
    :return: a float being the IRR.
    """

    if isinstance(values, Range):
        values = values.values

    if all(value < 0 for value in values):
        return 0

    if isinstance(dates, Range):
        dates = dates.values

    if is_not_number_input(values):
        return numeric_error(values, 'values')

    if is_not_number_input(dates):
        return numeric_error(dates, 'dates')

    if guess is not None and guess != 0:
        raise ValueError('guess value for excellib.irr() is %s and not 0' % guess)
    else:
        try:
            try:
                return scipy.optimize.newton(lambda r: xnpv(r, values, dates, lim_rate_low=False, lim_rate_high=True), 0.0)
            except (RuntimeError, FloatingPointError, ExcelError):  # Failed to converge?
                return scipy.optimize.brentq(lambda r: xnpv(r, values, dates, lim_rate_low=False, lim_rate_high=True), -1.0, 1e5)
        except Exception:
            return ExcelError('#NUM', 'IRR did not converge.')


def vlookup(lookup_value, table_array, col_index_num, range_lookup = True): # https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1

    if not isinstance(table_array, Range):
        return ExcelError('#VALUE', 'table_array should be a Range')

    if col_index_num > table_array.ncols:
        return ExcelError('#VALUE', 'col_index_num is greater than the number of cols in table_array')

    first_column = table_array.get(0, 1)
    result_column = table_array.get(0, col_index_num)

    if not range_lookup:
        if lookup_value not in first_column.values:
            return ExcelError('#N/A', 'lookup_value not in first column of table_array')
        else:
            i = first_column.values.index(lookup_value)
            ref = first_column.order[i]
    else:
        i = None
        for v in first_column.values:
            if lookup_value >= v:
                i = first_column.values.index(v)
                ref = first_column.order[i]
            else:
                break

        if i is None:
            return ExcelError('#N/A', 'lookup_value smaller than all values of table_array')

    return Range.find_associated_value(ref, result_column)


def sln(cost, salvage, life): # Excel reference: https://support.office.com/en-us/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8

    for arg in [cost, salvage, life]:
        if isinstance(arg, ExcelError) or arg in ErrorCodes:
            return arg

    return (cost - salvage) / life


def vdb(cost, salvage, life, start_period, end_period, factor = 2, no_switch = False): # Excel reference: https://support.office.com/en-us/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73

    for arg in [cost, salvage, life, start_period, end_period, factor, no_switch]:
        if isinstance(arg, ExcelError) or arg in ErrorCodes:
            return arg

    for arg in [cost, salvage, life, start_period, end_period, factor]:
        if not isinstance(arg, (float, int)):
            return ExcelError('#VALUE', 'Arg %s should be an int, float or long, instead: %s' % (arg, type(arg)))

    start_period = start_period
    end_period = end_period

    sln_depr = sln(cost, salvage, life)

    depr_rate = factor / life
    acc_depr = 0
    depr = 0

    switch_to_sln = False
    sln_depr = 0

    result = 0

    start_life = 0

    delta_life = life % 1
    if delta_life > 0: # to handle cases when life is not an integer
        end_life = int(life + 1)
    else:
        end_life = int(life)
    periods = list(range(start_life, end_life))

    if int(start_period) != start_period:
        delta_start = abs(int(start_period) - start_period)

        depr = (cost - acc_depr) * depr_rate * delta_start
        acc_depr += depr

        start_life = 1

        periods = [x + 0.5 for x in periods]

    for index, current_year in enumerate(periods):

        if not no_switch: # no_switch = False (Default Case)
            if switch_to_sln:
                depr = sln_depr
            else:
                depr = (cost - acc_depr) * depr_rate
                acc_depr += depr

                temp_sln_depr = sln(cost, salvage, life)

                if depr < temp_sln_depr:
                    switch_to_sln = True
                    fixed_remaining_years = life - current_year - 1
                    fixed_remaining_cost = cost - acc_depr

                     # we need to check future sln: current depr should never be smaller than sln to come
                    sln_depr = sln(fixed_remaining_cost, salvage, fixed_remaining_years)

                    if sln_depr > depr: # if it's the case, we switch to sln earlier than the regular case
                        # cancel what has been done
                        acc_depr -= depr
                        fixed_remaining_years += 1
                        fixed_remaining_cost = cost - acc_depr

                        # recalculate depreciation
                        sln_depr = sln(fixed_remaining_cost, salvage, fixed_remaining_years)
                        depr = sln_depr
                        acc_depr += depr
        else: # no_switch = True
            depr = (cost - acc_depr) * depr_rate
            acc_depr += depr

        delta_start = abs(current_year - start_period)

        if delta_start < 1 and delta_start != 0:
            result += depr * (1 - delta_start)
        elif current_year >= start_period and current_year < end_period:

            delta_end = abs(end_period - current_year)

            if delta_end < 1 and delta_end != 0:
                result += depr * delta_end
            else:
                result += depr

    return result


def xnpv(rate, values, dates, lim_rate_low=True, lim_rate_high=False):  # Excel reference: https://support.office.com/en-us/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7
    """
    Function to calculate the net present value (NPV) using payments and non-periodic dates. It resembles the excel function XPNV().

    :param rate: the discount rate.
    :param values: the payments of which at least one has to be negative.
    :param dates: the dates as excel dates (e.g. 43571 for 16/04/2019).
    :param lim_rate_low: to limit the rate below 0.
    :param lim_rate_high: to limit the rate above 1000 to avoid overflow errors.
    :return: a float being the NPV.
    """
    if isinstance(values, Range):
        values = values.values

    if isinstance(dates, Range):
        dates = dates.values

    if is_not_number_input(rate):
        return numeric_error(rate, 'rate')

    if is_not_number_input(values):
        return numeric_error(values, 'values')

    if is_not_number_input(dates):
        return numeric_error(dates, 'dates')

    if len(values) != len(dates):
        return ExcelError('#NUM!', '`values` range must be the same length as `dates` range in XNPV, %s != %s' % (len(values), len(dates)))

    if lim_rate_low and rate < 0:
        return ExcelError('#NUM!', '`excel cannot handle a negative `rate`' % (len(values), len(dates)))

    if lim_rate_high and rate > 1000:
        raise ExcelError('#NUM!', '`will result in an overflow error due to high `rate`')

    xnpv = 0
    with np.errstate(all='raise'):
        for v, d in zip(values, dates):
            xnpv += v / np.power(1.0 + rate, (d - dates[0]) / 365)

    return xnpv


def pmt(*args): # Excel reference: https://support.office.com/en-us/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441
    rate = args[0]
    num_payments = args[1]
    present_value = args[2]
    # WARNING fv & type not used yet - both are assumed to be their defaults (0)
    # fv = args[3]
    # type = args[4]
    return -present_value * rate / (1 - np.power(1 + rate, -num_payments))


# https://support.office.com/en-us/article/POWER-function-D3F2908B-56F4-4C3F-895A-07FB519C362A
def power(number, power):

    if number == power == 0:
        # Really excel?  What were you thinking?
        return ExcelError('#NUM!', 'Number and power cannot both be zero' % str(number))

    if power < 1 and number < 0:
        return ExcelError('#NUM!', '%s must be non-negative' % str(number))

    return np.power(number, power)


# https://support.office.com/en-ie/article/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf
def sqrt(number):
    if number < 0:
        return ExcelError('#NUM!', '%s must be non-negative' % str(index_num))
    return np.sqrt(number)


# https://support.office.com/en-ie/article/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9
def today():
    reference_date = datetime.datetime.today().date()
    days_since_epoch = reference_date - EXCEL_EPOCH
    # why +2 ?
    # 1 based from 1900-01-01
    # I think it is "inclusive" / to the _end_ of the day.
    # https://support.office.com/en-us/article/date-function-e36c0c8c-4104-49da-ab83-82328b832349
    """Note: Excel stores dates as sequential serial numbers so that they can be used in calculations.
    January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,447 days after January 1, 1900.
     You will need to change the number format (Format Cells) in order to display a proper date."""
    return days_since_epoch.days + 2


# https://support.office.com/en-us/article/concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2
def concat(*args):
    return concatenate(*tuple(flatten(args)))


# https://support.office.com/en-us/article/CONCATENATE-function-8F8AE884-2CA8-4F7A-B093-75D702BEA31D
# Important: In Excel 2016, Excel Mobile, and Excel Online, this function has
# been replaced with the CONCAT function. Although the CONCATENATE function is
# still available for backward compatibility, you should consider using CONCAT
# from now on. This is because CONCATENATE may not be available in future
# versions of Excel.
#
# BE AWARE; there are functional differences between CONACTENATE AND CONCAT
#
def concatenate(*args):
    if tuple(flatten(args)) != args:
        return ExcelError('#VALUE', 'Could not process arguments %s' % (args))

    cat_string = ''.join(str(a) for a in args)

    if len(cat_string) > CELL_CHARACTER_LIMIT:
        return ExcelError('#VALUE', 'Too long. concatentaed string should be no longer than %s but is %s' % (CELL_CHARACTER_LIMIT, len(cat_String)))

    return cat_string

if __name__ == '__main__':
    pass
