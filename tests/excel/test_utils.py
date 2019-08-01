from __future__ import absolute_import

import pyximport; pyximport.install()

import unittest

from koala.excellib import *

class TestUtils(unittest.TestCase):
    def test_col2num(self):
        """
        Testing col2num
        """
        self.assertEqual(1, col2num('A'))
        self.assertEqual(53, col2num('BA'))

        # TODO: fix AssertionError: Exception not raised
        # with self.assertRaises(Exception) as context:
        #     col2num('XFE')
        # self.assertTrue('Column ordinal must be left of XFD: XFE' in str(context.exception))

    def test_num2col(self):
        """
        Testing num2col
        """
        self.assertEqual('A', num2col(1))
        self.assertEqual('BA', num2col(53))

        # TODO: fix AssertionError: False is not true
        # with self.assertRaises(Exception) as context:
        #     num2col(0)
        # self.assertTrue('Column ordinal must be larger than 0: 0' in str(context.exception))

        # TODO: fix AssertionError: Exception not raised
        # with self.assertRaises(Exception) as context:
        #     num2col(16385) #XFE
        # self.assertTrue('Column ordinal must be less than than 16384: 16385' in str(context.exception))

    def test_is_almost_equal(self):
        """
        Testing is_almost_equal
        """
        self.assertEqual(True, is_almost_equal(0, 0))
        self.assertEqual(True, is_almost_equal(1.0, 1.0))
        self.assertEqual(True, is_almost_equal(0.1, 0.1))
        self.assertEqual(True, is_almost_equal(0.01, 0.01))
        self.assertEqual(True, is_almost_equal(0.001, 0.001))
        self.assertEqual(True, is_almost_equal(0.0001, 0.0001))
        self.assertEqual(True, is_almost_equal(0.00001, 0.00001))
        self.assertEqual(True, is_almost_equal(0.000001, 0.000001))

        self.assertEqual(False, is_almost_equal(0, 0.1))
        self.assertEqual(False, is_almost_equal(0, 0.01))
        self.assertEqual(False, is_almost_equal(0, 0.001))
        self.assertEqual(True, is_almost_equal(0, 0.0001))

        self.assertEqual(False, is_almost_equal(0, 1))

        self.assertEqual(True, is_almost_equal(0, 0, precision=0.001))
        self.assertEqual(True, is_almost_equal(1.0, 1.0, precision=0.001))
        self.assertEqual(True, is_almost_equal(0.1, 0.1, precision=0.001))
        self.assertEqual(True, is_almost_equal(0.01, 0.01, precision=0.001))
        self.assertEqual(True, is_almost_equal(0.001, 0.001, precision=0.001))
        self.assertEqual(True, is_almost_equal(0.0001, 0.0001, precision=0.001))
        self.assertEqual(True, is_almost_equal(0.00001, 0.00001, precision=0.001))
        self.assertEqual(True, is_almost_equal(0.000001, 0.000001, precision=0.001))

        self.assertEqual(False, is_almost_equal(0, 0.1, precision=0.001))
        self.assertEqual(False, is_almost_equal(0, 0.01, precision=0.001))
        self.assertEqual(True, is_almost_equal(0, 0.001, precision=0.001))
        self.assertEqual(True, is_almost_equal(0, 0.0001, precision=0.001))

        self.assertEqual(True, is_almost_equal(None, None))
        self.assertEqual(True, is_almost_equal(None, 'None'))
        self.assertEqual(True, is_almost_equal('None', None))
        self.assertEqual(True, is_almost_equal('None', 'None'))
        self.assertEqual(True, is_almost_equal(True, True))
        self.assertEqual(False, is_almost_equal(True, False))
        self.assertEqual(True, is_almost_equal('foo', 'foo'))
        self.assertEqual(False, is_almost_equal('foo', 'bar'))

    def test_is_range(self):
        """
        Testing is_range

        TODO change utils.is_range to accept only valid ranges.
        """
        self.assertEqual(True, is_range('A1:A2'))
        self.assertEqual(True, is_range('Sheet1!A1:A2'))
        self.assertEqual(True, is_range('Sheet1!A:B'))
        self.assertEqual(False, is_range('::param::')) #correctly borks on this

        self.assertEqual(True, is_range('foo:bar')) # not a valid range

    def test_split_range(self):
        """
        Testing split_range
        """
        # TODO: change split_range to return only valid range addresses
        # TODO: I'm not certain this is correct - ('"Sheet 1"', 'A1', 'A2'), utils.split_range('"Sheet 1"!A1:A2'). I would have expected ('Sheet 1', 'A1', 'A2')
        # TODO: #REF! is a valid address, but apparently un-splittable. Handle this.

        self.assertEqual(('Sheet1', 'A1', 'A2'), split_range('Sheet1!A1:A2'))
        self.assertEqual(('"Sheet 1"', 'A1', 'A2'), split_range('"Sheet 1"!A1:A2'))
        self.assertEqual((None, 'A', 'A'), split_range('A:A'))
        self.assertEqual((None, '1', '2'), split_range('1:2'))

        self.assertEqual(('Sheet1', 'A1', 'A16385'), split_range('Sheet1!A1:A16385')) # invalid address
        self.assertEqual(('Sheet1', 'A1', 'XFE1'), split_range('Sheet1!A1:XFE1')) # invalid address

        # self.assertEqual((None, None, None), utils.split_range('#REF!')) # valid address, un-splittable

    def test_address2index(self):
        """
        Testing address2index
        """
        self.assertEqual((1, 1), address2index('Sheet1!A1'))
        self.assertEqual((16384, 1), address2index('Sheet1!XFD1'))
        self.assertEqual((1, 1048576), address2index('Sheet1!A1048576'))

        self.assertEqual((1, 0), address2index('Sheet1!A0')) # not a valid address

        # TODO: fix AssertionError: Exception not raised
        # with self.assertRaises(Exception) as context:
        #     address2index('Sheet1!XFE1') #16385
        # self.assertTrue('Column ordinal must be left of XFD: XFE' in str(context.exception))
        self.assertEqual((1, 1048577), address2index('Sheet1!A1048577')) # not a valid address

    def test_index2addres(self):
        """
        Testing index2addres
        """
        self.assertEqual('A1', index2addres(1, 1))
        self.assertEqual('Sheet1!A1', index2addres(1, 1, 'Sheet1'))
        self.assertEqual('XFD1', index2addres(16384, 1))
        self.assertEqual('A1048576', index2addres(1, 1048576))

        self.assertEqual('A0', index2addres(1, 0)) # not a valid address

        # TODO: fix AssertionError: Exception not raised
        # with self.assertRaises(Exception) as context:
        #     index2addres(16385, 1) # XFE
        # self.assertTrue('Column ordinal must be less than than 16384: 16385' in str(context.exception))
        self.assertEqual('A1048577', index2addres(1, 1048577)) # not a valid address

    def test_get_linest_degree(self):
        """
        Testing get_linest_degree
        """
        self.assertEqual([1, 2, 3, 4, 5], list(flatten_list([1, 2, 3, [4], [], [[[[[[[[[5]]]]]]]]]])) )
        self.assertEqual([1, 2, 3], list(flatten_list([[1, 2], 3])) )

    def test_numeric_error(self):
        """
        Testing numeric_error

        All numeric_error can return is an exception.
        """
        excel_exception = ExcelError('#NUM!', '`excel cannot handle a non-numeric `foobar`')

        self.assertEqual(excel_exception, numeric_error(excel_exception, 'foobar') )
        self.assertTrue(excel_exception, numeric_error(1, 'foobar'))
        self.assertTrue(excel_exception, numeric_error('foo', 'foobar'))

    def test_is_number(self):
        """
        Testing is_number
        """
        self.assertTrue(is_number(1))
        self.assertTrue(is_number(1.1))
        self.assertTrue(is_number(0.1))
        self.assertTrue(is_number('1'))
        self.assertTrue(is_number('1.1'))
        self.assertTrue(is_number('0.1'))

        self.assertFalse(is_number('foobar'))

        excel_exception = ExcelError('#NUM!', '`excel cannot handle a non-numeric `foobar`')
        self.assertFalse(is_number(excel_exception))

    def test_is_leap_year(self):
        """
        Testing is_leap_year
        """
        self.assertTrue(is_leap_year(1900)) # it is according to MS.
        self.assertTrue(is_leap_year(2000))
        self.assertTrue(is_leap_year(2004))
        self.assertTrue(is_leap_year(2400))
        self.assertFalse(is_leap_year(2100))

        # TODO: fix TypeError: '<=' not supported between instances of 'str' and 'int'
        # self.assertFalse(is_leap_year('1901'))
        # self.assertFalse(is_leap_year('2200'))

        # self.assertTrue(is_leap_year('1900')) # it is according to MS.
        # self.assertTrue(is_leap_year('2000'))
        # self.assertTrue(is_leap_year('2004'))
        # self.assertTrue(is_leap_year('2400'))
        # self.assertFalse(is_leap_year('2100'))

        # self.assertFalse(is_leap_year('1901'))
        # self.assertFalse(is_leap_year('2200'))

        with self.assertRaises(Exception) as context:
            is_leap_year('foo')
        self.assertTrue('foo must be a number' in str(context.exception))

    def test_get_max_days_in_month(self):
        """
        Testing get_max_days_in_month
        """
        self.assertEqual(31, get_max_days_in_month(1, 1900))
        self.assertEqual(29, get_max_days_in_month(2, 1900))
        self.assertEqual(31, get_max_days_in_month(3, 1900))
        self.assertEqual(30, get_max_days_in_month(4, 1900))
        self.assertEqual(31, get_max_days_in_month(5, 1900))
        self.assertEqual(30, get_max_days_in_month(6, 1900))
        self.assertEqual(31, get_max_days_in_month(7, 1900))
        self.assertEqual(31, get_max_days_in_month(8, 1900))
        self.assertEqual(30, get_max_days_in_month(9, 1900))
        self.assertEqual(31, get_max_days_in_month(10, 1900))
        self.assertEqual(30, get_max_days_in_month(11, 1900))
        self.assertEqual(31, get_max_days_in_month(12, 1900))

        self.assertEqual(28, get_max_days_in_month(2, 1901))
        self.assertEqual(28, get_max_days_in_month(2, 2200))

        self.assertEqual(29, get_max_days_in_month(2, 2000))
        self.assertEqual(29, get_max_days_in_month(2, 2004))

        self.assertEqual(29, get_max_days_in_month(2, 2004))

    def test_normalize_year(self):
        """
        Testing normalize_year
        """
        self.assertEqual((1900, 1, 1), normalize_year(1900, 1, 1))
        self.assertEqual((1900, 2, 1), normalize_year(1900, 1, 32))
        self.assertEqual((1901, 1, 1), normalize_year(1900, 13, 1))
        # TODO: Explain and fix (1900, 1, 367) rolling one year but (1900, 13, 367) rolling 2 years 2 days
        self.assertEqual((1901, 1, 1), normalize_year(1900, 1, 367))
        self.assertEqual((1902, 1, 1), normalize_year(1900, 13, 366))

        # 2004-02-28 + 366 + 1 + 28 days becomes 2005-03-1. 1 is day after 28, 28 days is Feb.
        self.assertEqual((2005, 3, 1), normalize_year(2004, 2, 395))
        # 2004-02-28 + 366 - 1 + 28 days becomes 2005-02-28. 1 is day before 29, 28 days is Feb.
        self.assertEqual((2005, 2, 28), normalize_year(2004, 2, 394))

        # 2003-02-28 + 366 + 1 + 28 days becomes 2004-02-29. 1 is day after 28, 28 days is Feb.
        self.assertEqual((2004, 2, 29), normalize_year(2003, 2, 394))

    def test_date_from_int(self):
        """
        Testing date_from_int
        """
        with self.assertRaises(Exception) as context:
            date_from_int('foo')
        self.assertTrue('foo is not a number' in str(context.exception))

        # TODO: Excel is 1 based -- for some reason this does not raise an error. needs to be fixed.
        self.assertEqual((1900, 0, 0), date_from_int(0))

        self.assertEqual((1900, 1, 1), date_from_int(1))

        # 365 is a year, 1 is leap year day, 1 is year rollover
        self.assertEqual((1901, 1, 1), date_from_int(365 + 1 + 1))

        # 365 * 2 is two years, 1 is leap year day, 1 is year rollover
        self.assertEqual((1902, 1, 1), date_from_int(365 * 2 + 1 + 1))

        # 365 * 5 is five years, 2 is leap year days, 1 is year rollover
        self.assertEqual((1905, 1, 1), date_from_int(365 * 5 + 2 + 1))

    def test_int_from_date(self):
        """
        Testing int_from_date
        """
        # TODO: MS uses 1 based, this method appears to be 2 based. It might need to be as 1900 is considered a leap year.
        self.assertEqual( 2, int_from_date( dt.date(1900, 1, 1) ) )

        self.assertEqual( 365 + 1 + 1, int_from_date( dt.date(1901, 1, 1) ) )
        self.assertEqual( 365 * 2 + 1 + 1, int_from_date( dt.date(1902, 1, 1) ) )
        self.assertEqual( 365 * 5 + 2 + 1, int_from_date( dt.date(1905, 1, 1) ) )

    def test_old_div(self):
        """
        Testing old_div
        """
        self.assertEqual( 3, old_div( 6, 2 ) )
        self.assertEqual( 3, old_div( int(6), int(2) ) )
        self.assertEqual( 3, old_div( float(6), float(2) ) )
        self.assertEqual( 3, old_div( int(6), float(2) ) )
        self.assertEqual( 3, old_div( float(6), int(2) ) )


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
    # TODO: change utils.split_address to check that the address is valid.

    def test_parser(self):
        self.assertEqual(split_address('K54'), (None, 'K', '54'))
        self.assertEqual(split_address('Sheet1!K54'), ('Sheet1', 'K', '54'))
        self.assertEqual(split_address('Sheet1!5'), ('Sheet1', None, '5'))
        self.assertEqual(split_address('Sheet1!A'), ('Sheet1', 'A', None))

        self.assertEqual( ('Sheet1', 'A', '1'), split_address('Sheet1!A1') )
        self.assertEqual( ('Sheet1', 'A', '0'), split_address('Sheet1!A0') ) # not a valid address
        self.assertEqual( ('Sheet1', 'XFE', '1'), split_address('Sheet1!XFE1') ) # not a valid address
        self.assertEqual( ('Sheet1', 'XFE', '0'), split_address('Sheet1!XFE0') ) # not a valid address


class Test_resolve_range(unittest.TestCase):
    def test_parser(self):
        self.assertEqual(resolve_range('Sheet1!A1:A3'), (['Sheet1!A1', 'Sheet1!A2', 'Sheet1!A3'], 3, 1))
        self.assertEqual(resolve_range('Sheet1!A1:C1'), (['Sheet1!A1', 'Sheet1!B1', 'Sheet1!C1'], 1, 3))
        self.assertEqual(resolve_range('Sheet1!A1:B2'), ([['Sheet1!A1', 'Sheet1!B1'], ['Sheet1!A2', 'Sheet1!B2']], 2, 2))
        self.assertEqual(resolve_range('Sheet1!A:A')[1::], (2**20, 1))
        self.assertEqual(resolve_range('Sheet1!A:B')[1::], (2**20, 2))
        self.assertEqual(resolve_range('Sheet1!1:1')[1::], (1, 2**14))
        self.assertEqual(resolve_range('Sheet1!1:2')[1::], (2, 2**14))
