import unittest
import os
import sys

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../..')
sys.path.insert(0, path)

from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from koala.ast.graph import ExcelCompiler

from koala.ast.Range import Range, get_values


class Test_Excel(unittest.TestCase):
    
    def setUp(self):
        # This needs to be in setup so that further tests begin from scratch
        file_name = "./tests/ast/basic_evaluation.xlsx"

        c = ExcelCompiler(file_name)
        self.sp = c.gen_graph()
        
    def test_D1(self):
        self.sp.set_value('Sheet1!A1', 10)
    	self.assertEqual(self.sp.evaluate('Sheet1!D1'), 20)

    def test_D2(self):
        self.sp.set_value('Sheet1!A2', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!D2'), 30)

    def test_D3(self):
        self.sp.set_value('Sheet1!A3', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!D3'), 40)

    def test_E1(self):
        self.sp.set_value('Sheet1!B1', 20)
        self.assertEqual(self.sp.evaluate('Sheet1!E1'), 22)

    def test_F1(self):
        self.sp.set_value('Sheet1!B1', 20)
        self.assertEqual(self.sp.evaluate('Sheet1!F1'), 120)

    def test_G1(self):
        self.sp.set_value('Sheet1!B1', 20)
        self.assertEqual(self.sp.evaluate('Sheet1!G1'), 41)

    def test_D8(self):
        self.sp.set_value('Sheet1!A8', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!D8'), 17)

    def test_B6(self):
        self.sp.set_value('Sheet1!A6', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!B6'), 20)

    def test_J1(self):
        self.sp.set_value('Sheet1!B1', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!J1'), 4)



if __name__ == '__main__':
    unittest.main()