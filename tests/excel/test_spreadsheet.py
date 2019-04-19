import sys
import unittest

from koala.Spreadsheet import *
sys.setrecursionlimit(3000)

class Test_Spreadsheet(unittest.TestCase):
    def test_create(self):
        spreadsheet = Spreadsheet()

        spreadsheet.cell_add('Sheet1!A1', value=1)
        spreadsheet.cell_add('Sheet1!A2', value=2)
        spreadsheet.cell_add('Sheet1!A3', formula='=SUM(Sheet1!A1, Sheet1!A2)')
        spreadsheet.cell_add('Sheet1!A4', formula='=SUM(Sheet1!A1:A2)')
        spreadsheet.cell_add('Sheet1!A5', formula='=SUM(Sheet1!A1:A1)')

        self.assertEqual(spreadsheet.evaluate('Sheet1!A3'), 3)  # test function
        self.assertEqual(spreadsheet.evaluate('Sheet1!A4'), 3)  # test range
        self.assertEqual(spreadsheet.evaluate('Sheet1!A5'), 1)  # test short range
