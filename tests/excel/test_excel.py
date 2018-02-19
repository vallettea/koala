import os.path
import sys
import unittest

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../../')
sys.path.insert(0, path)

from koala.reader import read_archive, read_cells
from koala import ExcelCompiler


# # This fails, needs to be adapted
class Test_SharedFormula(unittest.TestCase):

    def setUp(self):
        file_name = os.path.abspath("./tests/files/SharedFormula.xlsx")
        archive = read_archive(file_name)
        self.cells = read_cells(archive)

    @unittest.skip('This test fails.')
    def test_nb_formulas(self):
        self.assertEqual(len([ref_cell for ref_cell in list(self.cells.items()) if ref_cell[1].formula is not None]), 13)

    def test_shared_formulas_content(self):
        self.assertEqual(self.cells[('Shared_formula!G2')].formula, 'G1 + 10 * L1 + $A$1')

    def test_text_content(self):
        self.assertEqual(self.cells[('Shared_formula!C12')].value, 'Romain')

    @unittest.skip('This test fails.')
    def test_types(self):
        nb_int = len([ref_cell1 for ref_cell1 in list(self.cells.items()) if type(ref_cell1[1].value) == int])
        nb_float = len([ref_cell2 for ref_cell2 in list(self.cells.items()) if type(ref_cell2[1].value) == float])
        nb_bool = len([ref_cell3 for ref_cell3 in list(self.cells.items()) if type(ref_cell3[1].value) == bool])
        nb_str = len([ref_cell4 for ref_cell4 in list(self.cells.items()) if type(ref_cell4[1].value) == str])

        self.assertTrue(nb_int == 21 and nb_float == 3 and nb_bool == 2 and nb_str == 10)


class Test_NamedRanges(unittest.TestCase):


    def setUp(self):
        c = ExcelCompiler("./tests/files/NamedRanges.xlsx", ignore_sheets = ['IHS'])
        self.graph = c.gen_graph()
        sys.setrecursionlimit(10000)

    def test_before_set_value(self):
        self.assertTrue(self.graph.evaluate('INPUT') == 1 and self.graph.evaluate('Sheet1!A1') == 1 and self.graph.evaluate('RESULT') == 187)

    @unittest.skip('This test fails.')
    def test_after_set_value(self):
        self.graph.set_value('INPUT', 2025)
        self.assertTrue(self.graph.evaluate('INPUT') == 2025 and self.graph.evaluate('Sheet1!A1') == 2025 and self.graph.evaluate('RESULT') == 2211)


if __name__ == '__main__':
    unittest.main()
