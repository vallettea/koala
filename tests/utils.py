from types import GeneratorType

import unittest
import datetime as dt
import numbers

import koala.utils as utils
from koala.ExcelError import ExcelError

class TestUtil(unittest.TestCase):

    def test_col2num(self):
        """
        Testing col2num
        """
        self.assertEqual(1, utils.col2num('A'))
        self.assertEqual(53, utils.col2num('BA'))

        with self.assertRaises(Exception) as context:
            utils.col2num('XFE')
        self.assertTrue('Column ordinal must be left of XFD: XFE' in str(context.exception))


    def test_num2col(self):
        """
        Testing num2col
        """
        self.assertEqual('A', utils.num2col(1))
        self.assertEqual('BA', utils.num2col(53))

        with self.assertRaises(Exception) as context:
            utils.num2col(0)
        self.assertTrue('Column ordinal must be larger than 0: 0' in str(context.exception))

        with self.assertRaises(Exception) as context:
            utils.num2col(16385) #XFE
        self.assertTrue('Column ordinal must be less than than 16384: 16385' in str(context.exception))

    def test_split_address(self):
        """
        Testing split_address

        TODO change utils.split_address to check that the address is valid.
        """
        self.assertEqual(('Sheet1', 'A', '1'), utils.split_address('Sheet1!A1'))
        self.assertEqual(('Sheet1', 'A', '0'), utils.split_address('Sheet1!A0')) # not a valid address
        self.assertEqual(('Sheet1', 'XFE', '1'), utils.split_address('Sheet1!XFE1')) # not a valid address
        self.assertEqual(('Sheet1', 'XFE', '0'), utils.split_address('Sheet1!XFE0')) # not a valid address

    def test_is_almost_equal(self):
        """
        Testing is_almost_equal
        """
        self.assertEqual(True, utils.is_almost_equal(0, 0))
        self.assertEqual(True, utils.is_almost_equal(1.0, 1.0))
        self.assertEqual(True, utils.is_almost_equal(0.1, 0.1))
        self.assertEqual(True, utils.is_almost_equal(0.01, 0.01))
        self.assertEqual(True, utils.is_almost_equal(0.001, 0.001))
        self.assertEqual(True, utils.is_almost_equal(0.0001, 0.0001))
        self.assertEqual(True, utils.is_almost_equal(0.00001, 0.00001))
        self.assertEqual(True, utils.is_almost_equal(0.000001, 0.000001))

        self.assertEqual(False, utils.is_almost_equal(0, 0.1))
        self.assertEqual(False, utils.is_almost_equal(0, 0.01))
        self.assertEqual(False, utils.is_almost_equal(0, 0.001))
        self.assertEqual(True, utils.is_almost_equal(0, 0.0001))

        self.assertEqual(False, utils.is_almost_equal(0, 1))

        self.assertEqual(True, utils.is_almost_equal(0, 0, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(1.0, 1.0, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0.1, 0.1, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0.01, 0.01, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0.001, 0.001, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0.0001, 0.0001, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0.00001, 0.00001, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0.000001, 0.000001, precision=0.001))

        self.assertEqual(False, utils.is_almost_equal(0, 0.1, precision=0.001))
        self.assertEqual(False, utils.is_almost_equal(0, 0.01, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0, 0.001, precision=0.001))
        self.assertEqual(True, utils.is_almost_equal(0, 0.0001, precision=0.001))

        self.assertEqual(True, utils.is_almost_equal(None, None))
        self.assertEqual(True, utils.is_almost_equal(None, 'None'))
        self.assertEqual(True, utils.is_almost_equal('None', None))
        self.assertEqual(True, utils.is_almost_equal('None', 'None'))
        self.assertEqual(True, utils.is_almost_equal(True, True))
        self.assertEqual(False, utils.is_almost_equal(True, False))
        self.assertEqual(True, utils.is_almost_equal('foo', 'foo'))
        self.assertEqual(False, utils.is_almost_equal('foo', 'bar'))

    def test_is_range(self):
        """
        Testing is_range

        TODO change utils.is_range to accept only valid ranges.
        """
        self.assertEqual(True, utils.is_range('A1:A2'))
        self.assertEqual(True, utils.is_range('Sheet1!A1:A2'))
        self.assertEqual(True, utils.is_range('Sheet1!A:B'))
        self.assertEqual(False, utils.is_range('::param::')) #correctly borks on this

        self.assertEqual(True, utils.is_range('foo:bar')) # not a valid range

    def test_split_range(self):
        """
        Testing split_range
        """
        # TODO: change split_range to return only valid range addresses
        # TODO: I'm not certain this is correct - ('"Sheet 1"', 'A1', 'A2'), utils.split_range('"Sheet 1"!A1:A2'). I would have expected ('Sheet 1', 'A1', 'A2')

        self.assertEqual(('Sheet1', 'A1', 'A2'), utils.split_range('Sheet1!A1:A2'))
        self.assertEqual(('"Sheet 1"', 'A1', 'A2'), utils.split_range('"Sheet 1"!A1:A2'))
        self.assertEqual((None, 'A', 'A'), utils.split_range('A:A'))
        self.assertEqual((None, '1', '2'), utils.split_range('1:2'))

        self.assertEqual(('Sheet1', 'A1', 'A16385'), utils.split_range('Sheet1!A1:A16385')) # invalid address
        self.assertEqual(('Sheet1', 'A1', 'XFE1'), utils.split_range('Sheet1!A1:XFE1')) # invalid address


        # self.assertEqual((None, None, None), utils.split_range('#REF!')) # valid address, un-splittable

    # def test_max_dimension(self):
    #     """
    #     Testing max_dimension
    #     """
    #     # need to know how to build or mock a cellmap, leaving this for
    #     pass

    # def test_resolve_range(self):
    #     """
    #     Testing resolve_range
    #     """
    #     # TODO: change split_range to return only valid range addresses
    #     # TODO: I'm not certain this is correct - ('"Sheet 1"', 'A1', 'A2'), utils.split_range('"Sheet 1"!A1:A2'). I would have expected ('Sheet 1', 'A1', 'A2')
    #
    #     pass

    def test_address2index(self):
        """
        Testing address2index
        """
        self.assertEqual((1, 1), utils.address2index('Sheet1!A1'))
        self.assertEqual((16384, 1), utils.address2index('Sheet1!XFD1'))
        self.assertEqual((1, 1048576), utils.address2index('Sheet1!A1048576'))

        self.assertEqual((1, 0), utils.address2index('Sheet1!A0')) # not a valid address

        with self.assertRaises(Exception) as context:
            utils.address2index('Sheet1!XFE1') #16385
        self.assertTrue('Column ordinal must be left of XFD: XFE' in str(context.exception))
        self.assertEqual((1, 1048577), utils.address2index('Sheet1!A1048577')) # not a valid address

    def test_index2addres(self):
        """
        Testing index2addres
        """
        self.assertEqual('A1', utils.index2addres(1, 1))
        self.assertEqual('Sheet1!A1', utils.index2addres(1, 1, 'Sheet1'))
        self.assertEqual('XFD1', utils.index2addres(16384, 1))
        self.assertEqual('A1048576', utils.index2addres(1, 1048576))

        self.assertEqual('A0', utils.index2addres(1, 0)) # not a valid address

        with self.assertRaises(Exception) as context:
            utils.index2addres(16385, 1) # XFE
        self.assertTrue('Column ordinal must be less than than 16384: 16385' in str(context.exception))
        self.assertEqual('A1048577', utils.index2addres(1, 1048577)) # not a valid address

    def test_get_linest_degree(self):
        """
        Testing get_linest_degree
        """
        self.assertEqual([1, 2, 3, 4, 5], list(utils.flatten_list([1, 2, 3, [4], [], [[[[[[[[[5]]]]]]]]]])) )
        self.assertEqual([1, 2, 3], list(utils.flatten_list([[1, 2], 3])) )

    def test_numeric_error(self):
        """
        Testing numeric_error

        All numeric_error can return is an exception.
        """
        excel_exception = ExcelError('#NUM!', '`excel cannot handle a non-numeric `foobar`')

        self.assertEqual(excel_exception, utils.numeric_error(excel_exception, 'foobar') )
        self.assertTrue(excel_exception, utils.numeric_error(1, 'foobar'))
        self.assertTrue(excel_exception, utils.numeric_error('foo', 'foobar'))


    def test_is_number(self):
        """
        Testing is_number
        """
        self.assertTrue(utils.is_number(1))
        self.assertTrue(utils.is_number(1.1))
        self.assertTrue(utils.is_number(0.1))
        self.assertTrue(utils.is_number('1'))
        self.assertTrue(utils.is_number('1.1'))
        self.assertTrue(utils.is_number('0.1'))

        self.assertFalse(utils.is_number('foobar'))

        excel_exception = ExcelError('#NUM!', '`excel cannot handle a non-numeric `foobar`')
        self.assertFalse(utils.is_number(excel_exception))

    def test_is_leap_year(self):
        """
        Testing is_leap_year
        """
        self.assertTrue(utils.is_leap_year(1900)) # it os according to MS.
        self.assertTrue(utils.is_leap_year(2000))
        self.assertTrue(utils.is_leap_year(2004))
        self.assertTrue(utils.is_leap_year(2400))
        self.assertFalse(utils.is_leap_year(2100))

        self.assertFalse(utils.is_leap_year('1901'))
        self.assertFalse(utils.is_leap_year('2200'))

        self.assertTrue(utils.is_leap_year('1900')) # it os according to MS.
        self.assertTrue(utils.is_leap_year('2000'))
        self.assertTrue(utils.is_leap_year('2004'))
        self.assertTrue(utils.is_leap_year('2400'))
        self.assertFalse(utils.is_leap_year('2100'))

        self.assertFalse(utils.is_leap_year('1901'))
        self.assertFalse(utils.is_leap_year('2200'))

        with self.assertRaises(Exception) as context:
            utils.is_leap_year('foo')
        self.assertTrue('foo must be a number' in str(context.exception))


    def test_get_max_days_in_month(self):
        """
        Testing get_max_days_in_month
        """
        self.assertEqual(31, utils.get_max_days_in_month(1, 1900))
        self.assertEqual(29, utils.get_max_days_in_month(2, 1900))
        self.assertEqual(31, utils.get_max_days_in_month(3, 1900))
        self.assertEqual(30, utils.get_max_days_in_month(4, 1900))
        self.assertEqual(31, utils.get_max_days_in_month(5, 1900))
        self.assertEqual(30, utils.get_max_days_in_month(6, 1900))
        self.assertEqual(31, utils.get_max_days_in_month(7, 1900))
        self.assertEqual(31, utils.get_max_days_in_month(8, 1900))
        self.assertEqual(30, utils.get_max_days_in_month(9, 1900))
        self.assertEqual(31, utils.get_max_days_in_month(10, 1900))
        self.assertEqual(30, utils.get_max_days_in_month(11, 1900))
        self.assertEqual(31, utils.get_max_days_in_month(12, 1900))

        self.assertEqual(28, utils.get_max_days_in_month(2, 1901))
        self.assertEqual(28, utils.get_max_days_in_month(2, 2200))

        self.assertEqual(29, utils.get_max_days_in_month(2, 2000))
        self.assertEqual(29, utils.get_max_days_in_month(2, 2004))

        self.assertEqual(29, utils.get_max_days_in_month(2, 2004))


    def test_normalize_year(self):
        """
        Testing normalize_year
        """
        self.assertEqual((1900, 1, 1), utils.normalize_year(1900, 1, 1))
        self.assertEqual((1900, 2, 1), utils.normalize_year(1900, 1, 32))
        self.assertEqual((1901, 1, 1), utils.normalize_year(1900, 13, 1))
        # TODO: Explain and fix (1900, 1, 367) rolling one year but (1900, 13, 367) rolling 2 years 2 days
        self.assertEqual((1901, 1, 1), utils.normalize_year(1900, 1, 367))
        self.assertEqual((1902, 1, 1), utils.normalize_year(1900, 13, 366))

        # 2004-02-28 + 366 + 1 + 28 days becomes 2005-03-1. 1 is day after 28, 28 days is Feb.
        self.assertEqual((2005, 3, 1), utils.normalize_year(2004, 2, 395))
        # 2004-02-28 + 366 - 1 + 28 days becomes 2005-02-28. 1 is day before 29, 28 days is Feb.
        self.assertEqual((2005, 2, 28), utils.normalize_year(2004, 2, 394))

        # 2003-02-28 + 366 + 1 + 28 days becomes 2004-02-29. 1 is day after 28, 28 days is Feb.
        self.assertEqual((2004, 2, 29), utils.normalize_year(2003, 2, 394))


    def test_date_from_int(self):
        """
        Testing date_from_int
        """
        with self.assertRaises(Exception) as context:
            utils.date_from_int('foo')
        self.assertTrue('foo is not a number' in str(context.exception))

        # TODO: Excel is 1 based -- for some reason this does not raise an error. needs to be fixed.
        self.assertEqual((1900, 0, 0), utils.date_from_int(0))

        self.assertEqual((1900, 1, 1), utils.date_from_int(1))

        # 365 is a year, 1 is leap year day, 1 is year rollover
        self.assertEqual((1901, 1, 1), utils.date_from_int(365 + 1 + 1))

        # 365 * 2 is two years, 1 is leap year day, 1 is year rollover
        self.assertEqual((1902, 1, 1), utils.date_from_int(365 * 2 + 1 + 1))

        # 365 * 5 is five years, 2 is leap year days, 1 is year rollover
        self.assertEqual((1905, 1, 1), utils.date_from_int(365 * 5 + 2 + 1))

    def test_int_from_date(self):
        """
        Testing int_from_date
        """
        # TODO: MS uses 1 based, this method appears to be 2 based. It might need to be as 1900 is considered a leap year.
        self.assertEqual( 2, utils.int_from_date( dt.date(1900, 1, 1) ) )

        self.assertEqual( 365 + 1 + 1, utils.int_from_date( dt.date(1901, 1, 1) ) )
        self.assertEqual( 365 * 2 + 1 + 1, utils.int_from_date( dt.date(1902, 1, 1) ) )
        self.assertEqual( 365 * 5 + 2 + 1, utils.int_from_date( dt.date(1905, 1, 1) ) )

    # def test_criteria_parser(self):
    #     """
    #     Testing criteria_parser
    #     """
    #     # TODO: This got too hard to test for now, will need to come back later. Seems specific to ExcelLib
    #     # check = utils.criteria_parser(2400)
    #     # valid = [index for index, item in enumerate(l) if check(item)]
    #     # self.assertFalse(valid[0])
    #     #
    #     # self.assertTrue(utils.criteria_parser('<'))

    # def test_find_corresponding_index(self):
    #     """
    #     Testing find_corresponding_index
    #     """
    #     # TODO: This got too hard to test for now, will need to come back later. Seems specific to ExcelLib
    #     # self.assertTrue(utils.find_corresponding_index((1, 2), "<"))

    # def test_check_length(self):
    #     """
    #     Testing check_length
    #     """
    #     # TODO: This got too hard to test for now, will need to come back later. Seems specific to ExcelLib

    # def test_extract_numeric_values(self):
    #     """
    #     Testing extract_numeric_values
    #     """
    #     # Seems specific to ExcelLib

    def test_old_div(self):
        """
        Testing old_div
        """
        self.assertEqual( 3, utils.old_div( 6, 2 ) )
        self.assertEqual( 3, utils.old_div( int(6), int(2) ) )
        self.assertEqual( 3, utils.old_div( float(6), float(2) ) )
        self.assertEqual( 3, utils.old_div( int(6), float(2) ) )
        self.assertEqual( 3, utils.old_div( float(6), int(2) ) )

    def test_safe_iterator(self):
        """
        Testing safe_iterator
        """



if __name__ == '__main__':
    unittest.main()
