from __future__ import absolute_import

import pyximport; pyximport.install()

import unittest

from koala.excellib import *


class Test_criteria_parser(unittest.TestCase):
    def test_parser_numeric(self):
        self.assertEqual(criteria_parser(2)(4), False)
        self.assertEqual(criteria_parser(3)(3), True)
        self.assertEqual(criteria_parser(4)(2), False)
        self.assertEqual(criteria_parser(4)('A'), False)
        self.assertEqual(criteria_parser(4)('4'), True)
        self.assertEqual(criteria_parser('4')(4), True)
        self.assertEqual(criteria_parser(4.0)('4'), True)
        self.assertEqual(criteria_parser('4')(4.0), True)

    def test_parser_not_equal_numeric(self):
        self.assertEqual(criteria_parser('<>3')(2), True)
        self.assertEqual(criteria_parser('<>3')(3), False)
        self.assertEqual(criteria_parser('<>3')(4), True)

    def test_parser_equal_numeric(self):
        self.assertEqual(criteria_parser('=3')(2), False)
        self.assertEqual(criteria_parser('=3')(3), True)
        self.assertEqual(criteria_parser('=3')(4), False)
        self.assertEqual(criteria_parser('=3.3')(3.3), True)
        self.assertEqual(criteria_parser('=3.0')(3), True)
        self.assertEqual(criteria_parser('=3')(3.0), True)

    def test_parser_smaller_than_numeric(self):
        self.assertEqual(criteria_parser('<3')(2), True)
        self.assertEqual(criteria_parser('<3')(3), False)
        self.assertEqual(criteria_parser('<3')(4), False)

    def test_parser_larger_than_numeric(self):
        self.assertEqual(criteria_parser('>3')(2), False)
        self.assertEqual(criteria_parser('>3')(3), False)
        self.assertEqual(criteria_parser('>3')(4), True)

    def test_parser_smaller_than_equal_numeric(self):
        self.assertEqual(criteria_parser('<=3')(2), True)
        self.assertEqual(criteria_parser('<=3')(3), True)
        self.assertEqual(criteria_parser('<=3')(4), False)

    def test_parser_larger_than_equal_numeric(self):
        self.assertEqual(criteria_parser('>=3')(2), False)
        self.assertEqual(criteria_parser('>=3')(3), True)
        self.assertEqual(criteria_parser('>=3')(4), True)

    def test_parser_strings(self):
        self.assertEqual(criteria_parser('A')('A'), True)
        self.assertEqual(criteria_parser('A')('a'), True)
        self.assertEqual(criteria_parser('a')('A'), True)
        self.assertEqual(criteria_parser('a')('a'), True)
        self.assertEqual(criteria_parser('A')('B'), False)
        self.assertEqual(criteria_parser('A')(1), False)

    def test_parser_strings_equality(self):
        self.assertEqual(criteria_parser('=A')('A'), True)
        self.assertEqual(criteria_parser('=A')('a'), True)
        self.assertEqual(criteria_parser('=a')('A'), True)
        self.assertEqual(criteria_parser('=a')('a'), True)
        self.assertEqual(criteria_parser('=A')('B'), False)
        self.assertEqual(criteria_parser('=A')(1), False)


class Test_split_address(unittest.TestCase):
    def test_parser(self):
        self.assertEqual(split_address('K54'), (None, 'K', '54'))
        self.assertEqual(split_address('Sheet1!K54'), ('Sheet1', 'K', '54'))
        self.assertEqual(split_address('Sheet1!5'), ('Sheet1', None, '5'))
        self.assertEqual(split_address('Sheet1!A'), ('Sheet1', 'A', None))


class Test_resolve_range(unittest.TestCase):
    def test_parser(self):
        self.assertEqual(resolve_range('Sheet1!A1:A3'), (['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3'], 3, 1))
        self.assertEqual(resolve_range('Sheet1!A1:C1'), (['Sheet1!A1', 'Sheet1!B1', 'Sheet1!C1'], 1, 3))
        self.assertEqual(resolve_range('Sheet1!A1:B2'), ([['Sheet1!A1', 'Sheet1!B1'], ['Sheet1!A2', 'Sheet1!B2']], 2, 2))
        self.assertEqual(resolve_range('Sheet1!A:A')[1::], (2**20, 1))
        self.assertEqual(resolve_range('Sheet1!A:B')[1::], (2**20, 2))
        self.assertEqual(resolve_range('Sheet1!1:1')[1::], (1, 2**14))
        self.assertEqual(resolve_range('Sheet1!1:2')[1::], (2, 2**14))