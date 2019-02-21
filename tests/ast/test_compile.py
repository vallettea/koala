# -*- coding: utf-8 -*-
import sys
import unittest

from koala import Cell, cell2code, RangeCore


class Test_cell2code(unittest.TestCase):
    def test_formula(self):
        cell = Cell(address="A1", sheet="sheet1")
        cell.formula = "10+20"

        code, ast = cell2code(cell, named_ranges=[])

        assert code == 'RangeCore.apply("add",10,20,(1, \'A\'))'

        RangeCore
        assert eval(code) == 30

    def test_formula_string(self):
        cell = Cell(address="A1", sheet="sheet1")
        cell.formula = "=IF(1, \"a\", \"b\")"

        code, ast = cell2code(cell, named_ranges=[])

        assert code == u'("a" if 1 else "b")'

        RangeCore
        assert eval(code) == u"a"

    @unittest.skipIf(sys.version_info <= (3, 0), "Unicode broken with Python 2")
    def test_formula_string_unicode(self):
        cell = Cell(address="A1", sheet="sheet1")
        cell.formula = "=IF(1, \"a☺\", \"b☺\")"

        code, ast = cell2code(cell, named_ranges=[])

        assert code == u'("a☺" if 1 else "b☺")'

        RangeCore
        assert eval(code) == u"a☺"

    def test_string(self):
        cell = Cell(address="A1", sheet="sheet1")
        cell.value = u"hello world"

        code, ast = cell2code(cell, named_ranges=[])

        assert code == u'u"hello world"'

        RangeCore
        assert eval(code) == u"hello world"

    def test_string_unicode(self):
        cell = Cell(address="A1", sheet="sheet1")
        cell.value = u"hello world ☺"

        code, ast = cell2code(cell, named_ranges=[])

        assert code == u'u"hello world ☺"'

        RangeCore
        assert eval(code) == u"hello world ☺"

    def test_string_quotes(self):
        cell = Cell(address="A1", sheet="sheet1")
        cell.value = u"hello \"world'"

        code, ast = cell2code(cell, named_ranges=[])

        assert code == u'u"hello \\"world\'"'

        RangeCore
        assert eval(code) == u"hello \"world'"
