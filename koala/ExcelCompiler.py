# cython: profile=True

from koala.Spreadsheet import Spreadsheet
import warnings


class ExcelCompiler(object):
    """Class responsible for taking cells and named_range and create a graph
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, file, ignore_sheets = [], ignore_hidden = False, debug = False):
        warnings.warn(
            "The ExcelCompiler class will disappear in a future version. Please use Spreadsheet instead.",
            PendingDeprecationWarning
        )
        self.spreadsheet = Spreadsheet(file = file, ignore_sheets = ignore_sheets, ignore_hidden = ignore_hidden, debug = debug)

    def clean_pointer(self):

        warnings.warn(
            "The ExcelCompiler class will disappear in a future version. Please use Spreadsheet.clean_pointer instead.",
            PendingDeprecationWarning
        )
        self.spreadsheet.clean_pointer()

    def gen_graph(self, outputs = [], inputs = []):

        warnings.warn(
            "The ExcelCompiler class will disappear in a future version. Please use Spreadsheet.gen_graph() instead. "
            "Please also note that this function is now included in the init of Spreadsheet and therefore it shouldn't "
            "be called as such anymore.",
            PendingDeprecationWarning
        )
        return self.spreadsheet
