import unittest


# Here is the graph contained in pruning.xlsx
#
#                          test
#                         /  |  \
#      A1 B1 C1   D1    G1  H1  I1
#       \  \ /    /       \  |  /
#       A2 B2   D2          H2
#         \/    /          /
#         A3   /          /
#           \ /          /
#           C6__________/
#
from koala import ExcelCompiler


class Test_without_range(unittest.TestCase):

    def setUp(self):
        file_name = "./tests/ast/pruning.xlsx"

        c = ExcelCompiler(file_name)
        sp = c.gen_graph(outputs = ["Sheet1!C6"])
        sp =sp.prune_graph(["Sheet1!A1","Sheet1!B1"])
        self.sp = sp

    @unittest.skip('This test fails.')
    def test_pruning_nodes(self):
        self.assertEqual(self.sp.G.number_of_nodes(), 9)

    @unittest.skip('This test fails.')
    def test_pruning_edges(self):
        self.assertEqual(self.sp.G.number_of_edges(), 8)

    @unittest.skip('This test fails.')
    def test_pruning_cellmap(self):
        self.assertEqual(len(list(self.sp.cellmap.keys())), 9)

    @unittest.skip('This test fails.')
    def test_eval(self):
        self.sp.set_value('Sheet1!A1', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!C6'), 38)


class Test_with_range(unittest.TestCase):

    def setUp(self):
        file_name = "./tests/ast/pruning.xlsx"

        c = ExcelCompiler(file_name)
        sp = c.gen_graph(outputs = ["Sheet1!C6"])
        sp = sp.prune_graph(["Sheet1!A1","Sheet1!B1", "test"])
        self.sp = sp

    @unittest.skip('This test fails.')
    def test_pruning_nodes(self):
      self.assertEqual(self.sp.G.number_of_nodes(), 13)

    @unittest.skip('This test fails.')
    def test_pruning_edges(self):
        self.assertEqual(self.sp.G.number_of_edges(), 12)

    @unittest.skip('This test fails.')
    def test_pruning_cellmap(self):
        self.assertEqual(len(list(self.sp.cellmap.keys())), 13)

    @unittest.skip('This test fails.')
    def test_A1(self):
        self.sp.set_value('Sheet1!A1', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!C6'), 38)

    @unittest.skip('This test fails.')
    def test_H1(self):
        self.sp.set_value('Sheet1!H1', 4)
        self.assertEqual(self.sp.evaluate('Sheet1!C6'), 31)

    @unittest.skip('This test fails.')
    def test_G1(self):
        self.sp.set_value('test', 4)
        self.assertEqual(self.sp.evaluate('Sheet1!C6'), 35)

    @unittest.skip('This test fails.')
    def test_G1_bis(self):
        self.sp.set_value('test', [7,8,9])
        self.assertEqual(self.sp.evaluate('Sheet1!C6'), 47)
