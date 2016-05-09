import unittest
import os
import sys
import json

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../..')
sys.path.insert(0, path)

from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from koala.ast.graph import ExcelCompiler

from koala.ast.Range import Range

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

    def test_J2(self):
        self.sp.set_value('Sheet1!B2', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!J2'), 0)

    def test_B17(self):
        self.sp.set_value('Sheet1!A17', 40)
        self.assertEqual(self.sp.evaluate('Sheet1!C17'), 80)

    def test_I17(self):
        self.sp.set_value('Sheet1!A2', 4)
        self.assertEqual(self.sp.evaluate('Sheet1!I17'), 4)

    def test_L1(self):
        self.sp.set_value('Sheet1!B1', 13)
        self.assertEqual(self.sp.evaluate('Sheet1!L1'), 13)

    def test_F26(self):
        self.sp.set_value('Sheet1!A23', 10)
        self.assertEqual(self.sp.evaluate('Sheet1!F26'), 21)

    def test_G26(self):
        self.sp.set_value('Sheet1!B22', 3)
        self.assertEqual(self.sp.evaluate('Sheet1!G26'), 10)

    def test_N1(self):
        self.sp.set_value('Sheet1!A1', 3)
        self.assertEqual(self.sp.evaluate('Sheet1!N1'), 3)

    def test_E32(self):
        self.sp.set_value('Sheet1!A31', 3)
        self.assertEqual(self.sp.evaluate('Sheet1!E32'), 19)

    def test_A37(self):
        self.sp.set_value('Sheet1!A36', 0.5)
        self.assertEqual(self.sp.evaluate('Sheet1!A37'), 0.52)

    def test_C37(self):
        self.sp.set_value('Sheet1!C36', 'David')
        self.assertEqual(self.sp.evaluate('Sheet1!C37'), 1)

    def test_G9(self):
        self.sp.set_value('Sheet1!A1', 2)
        self.assertEqual(self.sp.evaluate('Sheet1!G9'), 67)

    def test_P1(self):
        self.sp.set_value('Sheet1!A1', 2)
        self.assertEqual(self.sp.evaluate('Sheet1!P1'), 10)

    def test_Sheet2_B2(self):
        self.sp.set_value('Sheet1!B2',1000)
        self.assertEqual(self.sp.evaluate('Sheet2!B2'), 1000)

if __name__ == '__main__':
    unittest.main()