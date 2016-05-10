import unittest
import os
import sys
import json

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../..')
sys.path.insert(0, path)


from koala.ast.graph import ExcelCompiler


class Test_without_range(unittest.TestCase):
    
    def setUp(self):
        file_name = "./tests/ast/pruning.xlsx"

        c = ExcelCompiler(file_name)
        self.sp = c.gen_graph(outputs = ["Sheet1!C6"], inputs = ["Sheet1!A1","Sheet1!B1"])
        
    def test_pruning_nodes(self):
    	self.assertEqual(self.sp.G.number_of_nodes(), 9)

    def test_pruning_edges(self):
        self.assertEqual(self.sp.G.number_of_edges(), 8)

    def test_pruning_cellmap(self):
        self.assertEqual(len(self.sp.cellmap.keys()), 9)

class Test_with_range(unittest.TestCase):
    
    def setUp(self):
        file_name = "./tests/ast/pruning.xlsx"

        c = ExcelCompiler(file_name)
        self.sp = c.gen_graph(outputs = ["Sheet1!C6"], inputs = ["Sheet1!A1","Sheet1!B1", "test"])
        
    def test_pruning_nodes(self):
      self.assertEqual(self.sp.G.number_of_nodes(), 10)

    def test_pruning_edges(self):
        self.assertEqual(self.sp.G.number_of_edges(), 9)

    def test_pruning_cellmap(self):
        self.assertEqual(len(self.sp.cellmap.keys()), 10)

    def test_A1(self):
        self.sp.set_value('Sheet1!A1', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!C6'), 34)

    def test_H1(self):
        self.sp.set_value('Sheet1!H1', 4)
        self.assertEqual(self.sp.evaluate('Sheet1!C6'), 27)

    # def test_G1(self):
    #     self.sp.set_value('test', 4) # todo set_value for range
    #     print "zzzzz", self.sp.evaluate('Sheet1!C6')
    #     self.assertEqual(self.sp.evaluate('Sheet1!C6'), 27)

if __name__ == '__main__':
    unittest.main()


