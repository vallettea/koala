from __future__ import print_function
# cython: profile=True

import os.path

import networkx

from koala.reader import read_archive, read_named_ranges, read_cells
from koala.utils import *
from koala.ast import graph_from_seeds, shunting_yard, build_ast, prepare_pointer
from koala.Cell import Cell
from koala.Range import RangeFactory
from koala.Spreadsheet import Spreadsheet


class ExcelCompiler(object):
    """Class responsible for taking cells and named_range and create a graph
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, file, ignore_sheets = [], ignore_hidden = False, debug = False):
        # print("___### Initializing Excel Compiler ###___")

        self.spreadsheet = Spreadsheet(file = file, ignore_sheets = ignore_sheets, ignore_hidden = ignore_hidden, debug = debug)

    def clean_pointer(self):
        self.spreadsheet.clean_pointer()

    def gen_graph(self, outputs = [], inputs = []):
        self.spreadsheet

        return self.spreadsheet
