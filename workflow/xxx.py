import unittest

from koala.ExcelCompiler import ExcelCompiler
from koala.Cell import Cell

file_name = "../tests/ast/basic_evaluation.xlsx"

c = ExcelCompiler(file_name, debug = True)
sp = c.gen_graph()

sp.set_value('Sheet1!B2', 10)
sp.evaluate('Sheet1!J2')