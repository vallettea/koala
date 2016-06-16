import datetime
import re


class CellCoordinatesException(Exception):
    """Error for converting between numeric and A1-style cell references."""

FORMULAE = ("CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY", "CUBERANKEDMEMBER", "CUBESET", "CUBESETCOUNT", "CUBEVALUE", "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP", "DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS360", "EDATE", "EOMONTH", "HOUR", "MINUTE", "MONTH", "NETWORKDAYS", "NETWORKDAYS.INTL", "NOW", "SECOND", "TIME", "TIMEVALUE", "TODAY", "WEEKDAY", "WEEKNUM", "WORKDAY ", "WORKDAY.INTL", "YEAR", "YEARFRAC", "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "BIN2DEC", "BIN2HEX", "BIN2OCT", "COMPLEX", "CONVERT", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DELTA", "ERF", "ERFC", "GESTEP", "HEX2BIN", "HEX2DEC", "HEX2OCT", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSIN", "IMSQRT", "IMSUB", "IMSUM", "OCT2BIN", "OCT2DEC", "OCT2HEX", "ACCRINT", "ACCRINTM", "AMORDEGRC", "AMORLINC", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "CUMIPMT", "CUMPRINC", "DB", "DDB", "DISC", "DOLLARDE", "DOLLARFR", "DURATION", "EFFECT", "FV", "FVSCHEDULE", "INTRATE", "IPMT", "IRR", "ISPMT", "MDURATION", "MIRR", "NOMINAL", "NPER", "NPV", "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD", "PMT", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT", "PV", "RATE", "RECEIVED", "SLN", "SYD", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "VDB", "XIRR", "XNPV", "YIELD", "YIELDDISC", "YIELDMAT", "CELL", "ERROR.TYPE", "INFO", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT", "N", "NA", "TYPE", "AND", "FALSE", "IF", "IFERROR", "NOT", "OR", "TRUE ADDRESS", "AREAS", "CHOOSE", "COLUMN", "COLUMNS", "GETPIVOTDATA", "HLOOKUP", "HYPERLINK", "INDEX", "INDIRECT", "LOOKUP", "MATCH", "OFFSET", "ROW", "ROWS", "RTD", "TRANSPOSE", "VLOOKUP", "ABS", "ACOS", "ACOSH", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "CEILING", "COMBIN", "COS", "COSH", "DEGREES", "ECMA.CEILING", "EVEN", "EXP", "FACT", "FACTDOUBLE", "FLOOR", "GCD", "INT", "ISO.CEILING", "LCM", "LN", "LOG", "LOG10", "MDETERM", "MINVERSE", "MMULT", "MOD", "MROUND", "MULTINOMIAL", "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "RADIANS", "RAND", "RANDBETWEEN", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP", "SERIESSUM", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "TAN", "TANH", "TRUNC", "AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BETADIST", "BETAINV", "BINOMDIST", "CHIDIST", "CHIINV", "CHITEST", "CONFIDENCE", "CORREL", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "COVAR", "CRITBINOM", "DEVSQ", "EXPONDIST", "FDIST", "FINV", "FISHER", "FISHERINV", "FORECAST", "FREQUENCY", "FTEST", "GAMMADIST", "GAMMAINV", "GAMMALN", "GEOMEAN", "GROWTH", "HARMEAN", "HYPGEOMDIST", "INTERCEPT", "KURT", "LARGE", "LINEST", "LOGEST", "LOGINV", "LOGNORMDIST", "MAX", "MAXA", "MEDIAN", "MIN", "MINA", "MODE", "NEGBINOMDIST", "NORMDIST", "NORMINV", "NORMSDIST", "NORMSINV", "PEARSON", "PERCENTILE", "PERCENTRANK", "PERMUT", "POISSON", "PROB", "QUARTILE", "RANK", "RSQ", "SKEW", "SLOPE", "SMALL", "STANDARDIZE", "STDEV STDEVA", "STDEVP", "STDEVPA STEYX", "TDIST", "TINV", "TREND", "TRIMMEAN", "TTEST", "VAR", "VARA", "VARP", "VARPA", "WEIBULL", "ZTEST", "ASC", "BAHTTEXT", "CHAR", "CLEAN", "CODE", "CONCATENATE", "DOLLAR", "EXACT", "FIND", "FINDB", "FIXED", "JIS", "LEFT", "LEFTB", "LEN", "LENB", "LOWER", "MID", "MIDB", "PHONETIC", "PROPER", "REPLACE", "REPLACEB", "REPT", "RIGHT", "RIGHTB", "SEARCH", "SEARCHB", "SUBSTITUTE", "T", "TEXT", "TRIM", "UPPER", "VALUE")

FORMULAE = frozenset(FORMULAE)

# constants
COORD_RE = re.compile('^[$]?([A-Z]+)[$]?(\d+)$')
RANGE_EXPR = """
[$]?(?P<min_col>[A-Z]+)
[$]?(?P<min_row>\d+)
(:[$]?(?P<max_col>[A-Z]+)
[$]?(?P<max_row>\d+))?
"""
ABSOLUTE_RE = re.compile('^' + RANGE_EXPR +'$', re.VERBOSE)
SHEETRANGE_RE = re.compile("""
^(('(?P<quoted>([^']|'')*)')|(?P<notquoted>[^']*))!
(?P<cells>{0})$""".format(RANGE_EXPR), re.VERBOSE)

FLOAT_REGEX = re.compile(r"\.|[E-e]")

def _cast_number(value): # https://bitbucket.org/openpyxl/openpyxl/src/93604327bce7aac5e8270674579af76d390e09c0/openpyxl/cell/read_only.py?at=default&fileviewer=file-view-default
    "Convert numbers as string to an int or float"
    m = FLOAT_REGEX.search(value)
    if m is not None:
        return float(value)
    return int(value) # if no . nor E|e is found, it's an integer

def get_column_interval(start, end):
    start = column_index_from_string(start)
    end = column_index_from_string(end)
    return [get_column_letter(x) for x in range(start, end + 1)]


def coordinate_from_string(coord_string):
    """Convert a coordinate string like 'B12' to a tuple ('B', 12)"""
    match = COORD_RE.match(coord_string.upper())
    if not match:
        msg = 'Invalid cell coordinates (%s)' % coord_string
        raise CellCoordinatesException(msg)
    column, row = match.groups()
    row = int(row)
    if not row:
        msg = "There is no row 0 (%s)" % coord_string
        raise CellCoordinatesException(msg)
    return (column, row)


def absolute_coordinate(coord_string):
    """Convert a coordinate to an absolute coordinate string (B12 -> $B$12)"""
    m = ABSOLUTE_RE.match(coord_string.upper())
    if m:
        parts = m.groups()
        if all(parts[-2:]):
            return '$%s$%s:$%s$%s' % (parts[0], parts[1], parts[3], parts[4])
        else:
            return '$%s$%s' % (parts[0], parts[1])
    else:
        return coord_string

def _get_column_letter(col_idx):
    """Convert a column number into a column letter (3 -> 'C')

    Right shift the column col_idx by 26 to find column letters in reverse
    order.  These numbers are 1-based, and can be converted to ASCII
    ordinals by adding 64.

    """
    # these indicies correspond to A -> ZZZ and include all allowed
    # columns
    if not 1 <= col_idx <= 18278:
        raise ValueError("Invalid column index {0}".format(col_idx))
    letters = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx, 26)
        # check for exact division and borrow if needed
        if remainder == 0:
            remainder = 26
            col_idx -= 1
        letters.append(chr(remainder+64))
    return ''.join(reversed(letters))


_COL_STRING_CACHE = {}
_STRING_COL_CACHE = {}
for i in range(1, 18279):
    col = _get_column_letter(i)
    _STRING_COL_CACHE[i] = col
    _COL_STRING_CACHE[col] = i


def get_column_letter(idx,):
    """Convert a column index into a column letter
    (3 -> 'C')
    """
    try:
        return _STRING_COL_CACHE[idx]
    except KeyError:
        raise ValueError("Invalid column index {0}".format(idx))


def column_index_from_string(str_col):
    """Convert a column name into a numerical index
    ('A' -> 1)
    """
    # we use a function argument to get indexed name lookup
    try:
        return _COL_STRING_CACHE[str_col.upper()]
    except KeyError:
        raise ValueError("{0} is not a valid column name".format(str_col))


def range_boundaries(range_string):
    """
    Convert a range string into a tuple of boundaries:
    (min_col, min_row, max_col, max_row)
    Cell coordinates will be converted into a range with the cell at both end
    """
    m = ABSOLUTE_RE.match(range_string)
    min_col, min_row, sep, max_col, max_row = m.groups()
    min_col = column_index_from_string(min_col)
    min_row = int(min_row)

    if max_col is None or max_row is None:
        max_col = min_col
        max_row = min_row
    else:
        max_col = column_index_from_string(max_col)
        max_row = int(max_row)

    return min_col, min_row, max_col, max_row


def rows_from_range(range_string):
    """
    Get individual addresses for every cell in a range.
    Yields one row at a time.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    for row in range(min_row, max_row+1):
        yield tuple('%s%d' % (get_column_letter(col), row)
                    for col in range(min_col, max_col+1))


def cols_from_range(range_string):
    """
    Get individual addresses for every cell in a range.
    Yields one row at a time.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    for col in range(min_col, max_col+1):
        yield tuple('%s%d' % (get_column_letter(col), row)
                    for row in range(min_row, max_row+1))


def coordinate_to_tuple(coordinate):
    """
    Convert an Excel style coordinate to (row, colum) tuple
    """
    col, row = coordinate_from_string(coordinate)
    return row, _COL_STRING_CACHE[col]


def range_to_tuple(range_string):
    """
    Convert a worksheet range to the sheetname and maximum and minimum
    coordinate indices
    """
    m = SHEETRANGE_RE.match(range_string)
    if m is None:
        raise ValueError("Value must be of the form sheetname!A1:E4")
    sheetname = m.group("quoted") or m.group("notquoted")
    cells = m.group("cells")
    boundaries = range_boundaries(cells)
    return sheetname, boundaries