
ErrorCodes = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")

class ExcelError(Exception):
    pass

class EmptyCellError(ExcelError):
    pass