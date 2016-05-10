import unittest
import os
import sys
import json

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../..')
sys.path.insert(0, path)


from koala.ast.graph import ExcelCompiler


class Test_Excel(unittest.TestCase):
    
    def setUp(self):
        # This needs to be in setup so that further tests begin from scratch
        file_name = "./tests/ast/pruning.xlsx"

        c = ExcelCompiler(file_name)
        self.sp = c.gen_graph(outputs = ["Sheet1!C6"], inputs = ["Sheet1!A1","Sheet1!B1"])
        
    def test_pruning_nodes(self):
    	self.assertEqual(self.sp.G.number_of_nodes(), 8)

    def test_pruning_edges(self):
        self.assertEqual(self.sp.G.number_of_edges(), 7)

    def test_pruning_cellmap(self):
        self.assertEqual(len(self.sp.cellmap.keys()), 8)


if __name__ == '__main__':
    unittest.main()