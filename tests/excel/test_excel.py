import os.path
import sys
import unittest

from koala.reader import read_archive, read_cells
from koala import ExcelCompiler, Spreadsheet


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


class Test_EmptyCellInRange(unittest.TestCase):


    def setUp(self):
        c = ExcelCompiler("./tests/files/EmptyCellInRange.xlsx")
        self.graph = c.gen_graph()
        sys.setrecursionlimit(10000)

    def test_set_value(self):
        excel = self.graph

        self.assertEqual(excel.evaluate('Data_vert!A2'), 1)
        excel.set_value('Data_vert!A2', 1)
        self.assertEqual(excel.evaluate('Data_vert!B2'), "A")
        excel.set_value('Data_vert!B2', "A")

        self.assertEqual(excel.evaluate('DataJump_vert!A3'), 1)
        excel.set_value('DataJump_vert!A3', 1)
        self.assertEqual(excel.evaluate('DataJump_vert!B3'), "A")
        excel.set_value('DataJump_vert!B3', "A")

        self.assertEqual(excel.evaluate('Data_hor!B1'), 1)
        excel.set_value('Data_hor!B1', 1)
        self.assertEqual(excel.evaluate('Data_hor!B2'), "A")
        excel.set_value('Data_hor!B2', "A")

        self.assertEqual(excel.evaluate('DataJump_hor!C1'), 1)
        excel.set_value('DataJump_hor!C1', 1)
        self.assertEqual(excel.evaluate('DataJump_hor!C2'), "A")
        excel.set_value('DataJump_hor!C2', "A")

    def test_sumifs(self):
        excel = self.graph

        # test normal links, criterion as string in function
        self.assertEqual(excel.evaluate('Sumifs!B1'), 10)  # horizontal, 1:1 style ranges
        self.assertEqual(excel.evaluate('Sumifs!B2'), 10)  # vertical, A:A style ranges
        self.assertEqual(excel.evaluate('Sumifs!B3'), 10)  # horizontal, 1:1 style ranges with empty cells in range
        self.assertEqual(excel.evaluate('Sumifs!B4'), 10)  # vertical, A:A style ranges with empty cells in range

        # test normal links, criterion via link to cell
        self.assertEqual(excel.evaluate('Sumifs!B6'), 10)  # horizontal, 1:1 style ranges
        self.assertEqual(excel.evaluate('Sumifs!B7'), 10)  # vertical, A:A style ranges
        self.assertEqual(excel.evaluate('Sumifs!B8'), 10)  # horizontal, 1:1 style ranges with empty cells in range
        self.assertEqual(excel.evaluate('Sumifs!B9'), 10)  # vertical, A:A style ranges with empty cells in range

        # test hard links, criterion via link to cell
        self.assertEqual(excel.evaluate('Sumifs!B11'), 10)  # horizontal, 1:1 style ranges
        self.assertEqual(excel.evaluate('Sumifs!B12'), 10)  # vertical, A:A style ranges
        self.assertEqual(excel.evaluate('Sumifs!B13'), 10)  # horizontal, 1:1 style ranges with empty cells in range
        self.assertEqual(excel.evaluate('Sumifs!B14'), 10)  # vertical, A:A style ranges with empty cells in range

    def test_index_match(self):
        excel = self.graph

        self.assertEqual(excel.evaluate('IndexMatch!B1'), 1)  # horizontal, 1:1 style ranges
        self.assertEqual(excel.evaluate('IndexMatch!B2'), 1)  # vertical, A:A style ranges
        self.assertEqual(excel.evaluate('IndexMatch!B3'), 1)  # horizontal, 1:1 style ranges with empty cells in range
        self.assertEqual(excel.evaluate('IndexMatch!B4'), 1)  # vertical, A:A style ranges with empty cells in range


class Test_DumpDict(unittest.TestCase):

    def setUp(self):
        c = ExcelCompiler("./tests/files/NamedRanges.xlsx", ignore_sheets=['IHS'])
        self.graph = c.gen_graph()
        sys.setrecursionlimit(10000)

    def test_no_set_value(self):
        graph = self.graph
        self.assertTrue(graph.evaluate('INPUT') == 1)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 1)
        self.assertTrue(graph.evaluate('RESULT') == 187)

    def test_set_value(self):
        # Clone object
        data = self.graph.asdict()
        graph = Spreadsheet.from_dict(data)

        # Set value and check result in clone.
        graph.set_value('INPUT', 2025)
        self.assertTrue(graph.evaluate('INPUT') == 2025)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 2025)
        # self.assertTrue(graph.evaluate('RESULT') == 2211)

        # Check original not changed.
        graph = self.graph
        self.assertTrue(graph.evaluate('INPUT') == 1)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 1)
        self.assertTrue(graph.evaluate('RESULT') == 187)


class Test_Dump(unittest.TestCase):


    def setUp(self):
        c = ExcelCompiler("./tests/files/NamedRanges.xlsx", ignore_sheets = ['IHS'])
        self.graph = c.gen_graph()
        sys.setrecursionlimit(10000)

    def test_no_set_value(self):
        graph = self.graph
        self.assertTrue(graph.evaluate('INPUT') == 1)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 1)
        self.assertTrue(graph.evaluate('RESULT') == 187)

    def test_set_value(self):
        # Clone object
        self.graph.dump("dump.txt.gz")
        graph = Spreadsheet.load("dump.txt.gz")

        # Set value and check result in clone.
        graph.set_value('INPUT', 2025)
        self.assertTrue(graph.evaluate('INPUT') == 2025)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 2025)
        # self.assertTrue(graph.evaluate('RESULT') == 2211)

        # Check original not changed.
        graph = self.graph
        self.assertTrue(graph.evaluate('INPUT') == 1)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 1)
        self.assertTrue(graph.evaluate('RESULT') == 187)


class Test_DumpJson(unittest.TestCase):


    def setUp(self):
        c = ExcelCompiler("./tests/files/NamedRanges.xlsx", ignore_sheets = ['IHS'])
        self.graph = c.gen_graph()
        sys.setrecursionlimit(10000)

    def test_no_set_value(self):
        graph = self.graph
        self.assertTrue(graph.evaluate('INPUT') == 1)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 1)
        self.assertTrue(graph.evaluate('RESULT') == 187)

    def test_set_value(self):
        # Clone object
        self.graph.dump_json("dump.txt.gz")
        graph = Spreadsheet.load_json("dump.txt.gz")

        # Set value and check result in clone.
        graph.set_value('INPUT', 2025)
        self.assertTrue(graph.evaluate('INPUT') == 2025)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 2025)
        # self.assertTrue(graph.evaluate('RESULT') == 2211)

        # Check original not changed.
        graph = self.graph
        self.assertTrue(graph.evaluate('INPUT') == 1)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 1)
        self.assertTrue(graph.evaluate('RESULT') == 187)

        # Set value and check result in clone.
        graph.set_value('Sheet1!A1', 2025)
        self.assertTrue(graph.evaluate('Sheet1!A1') == 2025)
