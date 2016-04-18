import unittest
import glob

import os.path
import sys
from datetime import datetime

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../../')
sys.path.insert(0, path)

from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from koala.ast.graph import ExcelCompiler
from koala.ast.excelutils import Cell

class Test_SharedFormula(unittest.TestCase):
    
    def setUp(self):
        file_name = os.path.abspath("./tests/files/SharedFormula.xlsx")
        archive = read_archive(file_name)        
        self.cells = read_cells(archive)
        
    def test_nb_formulas(self):
        self.assertEqual(len(filter(lambda (ref, cell): cell.formula is not None, self.cells.items())), 13)

    def test_shared_formulas_content(self):
        self.assertEqual(self.cells[('Shared_formula!G2')].formula, 'G1 + 10 * L1 + $A$1')

    def test_text_content(self):
        self.assertEqual(self.cells[('Shared_formula!C12')].value, 'Romain')

    def test_types(self):
        nb_int = len(filter(lambda (ref, cell): type(cell.value) == int, self.cells.items()))
        nb_float = len(filter(lambda (ref, cell): type(cell.value) == float, self.cells.items()))
        nb_bool = len(filter(lambda (ref, cell): type(cell.value) == bool, self.cells.items()))
        nb_str = len(filter(lambda (ref, cell): type(cell.value) == str, self.cells.items()))

        self.assertTrue(nb_int == 21 and nb_float == 3 and nb_bool == 2 and nb_str == 10)


class Test_NamedRanges(unittest.TestCase):
    

    def setUp(self):
        c = ExcelCompiler("./tests/files/NamedRanges.xlsx", ignore_sheets = ['IHS'])
        self.graph = c.gen_graph()
        sys.setrecursionlimit(10000)
        
    def test_before_set_value(self):
        self.assertTrue(self.graph.evaluate('INPUT') == 1 and self.graph.evaluate('Sheet1!A1') == 1 and self.graph.evaluate('RESULT') == 187)

    def test_after_set_value(self):
        self.graph.set_value('INPUT', 2025)
        self.assertTrue(self.graph.evaluate('INPUT') == 2025 and self.graph.evaluate('Sheet1!A1') == 2025 and self.graph.evaluate('RESULT') == 2211)


if __name__ == '__main__':
    unittest.main()