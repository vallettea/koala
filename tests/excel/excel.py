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


class Test_Excel(unittest.TestCase):
    

    def setUp(self):
        file_name = os.path.abspath("../example/example2.xlsx")
        archive = read_archive(file_name)        
        self.cells = read_cells(archive) 

    def test_nb_formulas(self):
        self.assertEqual(len(filter(lambda (ref, cell): cell.formula is not None, self.cells.items())), 8)

    def test_shared_formulas_content(self):
        self.assertEqual(self.cells[('Shared_formula', 'G2')].formula, 'G1 + 10 * L1 + $A$1')

if __name__ == '__main__':
    unittest.main()