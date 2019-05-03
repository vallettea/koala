from __future__ import absolute_import

import pyximport; pyximport.install()

import unittest

from koala.excellib import *
from koala.ExcelError import ExcelError
from koala.Range import RangeCore


# https://support.office.com/en-ie/article/power-function-d3f2908b-56f4-4c3f-895a-07fb519c362a
class Test_Basics(unittest.TestCase):
    def test_divide(self):
        self.assertEqual(RangeCore.divide(2, 2), 1)
        self.assertIsInstance(RangeCore.divide(1, 0), ExcelError)
        self.assertIsInstance(RangeCore.divide(ExcelError('#VALUE'), 1), ExcelError)
        self.assertIsInstance(RangeCore.divide(1, ExcelError('#VALUE')), ExcelError)

    def test_multiply(self):
        self.assertEqual(RangeCore.multiply(2, 2), 4)
        self.assertEqual(RangeCore.multiply(1, 0), 0)
        self.assertIsInstance(RangeCore.multiply(ExcelError('#VALUE'), 1), ExcelError)
        self.assertIsInstance(RangeCore.multiply(1, ExcelError('#VALUE')), ExcelError)

    def test_power(self):
        self.assertEqual(RangeCore.power(2, 2), 4)
        self.assertIsInstance(RangeCore.power(ExcelError('#VALUE'), 1), ExcelError)
        self.assertIsInstance(RangeCore.power(1, ExcelError('#VALUE')), ExcelError)


class Test_VDB(unittest.TestCase):
    def test_vdb_basic(self):
        cost = 575000
        salvage = 5000
        life = 10
        rate = 1.5
        start = 3
        end = 5

        obj = 102160.546875

        self.assertEqual(vdb(cost, salvage, life, start, end, rate), obj)

    def test_vdb_partial(self):
        cost = 1
        salvage = 0
        life = 14
        rate = 1.25
        start = 11.5
        end = 12.5

        obj = 0.068726290454684

        self.assertEqual(round(vdb(cost, salvage, life, start, end, rate), 15), obj)

    def test_vdb_partial_no_switch(self):
        cost = 1
        salvage = 0
        life = 5.0
        rate = 2.5
        start = 0.5
        end = 1.5

        obj = 0.375

        self.assertEqual(vdb(cost, salvage, life, start, end, rate, True), obj)


class Test_SLN(unittest.TestCase):
    def test_sln_basic(self):
        self.assertEqual(sln(30000, 5000, 10), 2500)


class Test_Choose(unittest.TestCase):
    def test_choose_basic(self):
        self.assertEqual(choose(3, 'John', 'Paul', 'George', 'Ringo'), 'George')

    def test_choose_fraction(self):
        self.assertEqual(choose(3.4, 'John', 'Paul', 'George', 'Ringo'), 'George')

    def test_choose_incorrect_index(self):
        self.assertEqual(type(choose(3, 2)), ExcelError)


class Test_Irr(unittest.TestCase):
    def test_irr_errors(self):
        self.assertIsInstance(irr([-100, 39, 59, 55, ExcelError('#NUM')], 0), ExcelError)

    def test_irr_basic(self):
        self.assertEqual(round(irr([-100, 39, 59, 55, 20], 0), 7), 0.2809484)

    def test_irr_with_guess_non_null(self):
        with self.assertRaises(ValueError):
            irr([-100, 39, 59, 55, 20], 2)


class Test_Xirr(unittest.TestCase):
    def test_xirr_errors(self):
        self.assertIsInstance(xirr([-100, 30, 30, 30, ExcelError('#NUM')], [43571, 43721, 43871, 44021, 44171], 0), ExcelError)
        self.assertIsInstance(xirr([-100, 30, 30, 30, 30], [43571, 43721, 43871, 44021, ExcelError('#NUM')], 0), ExcelError)


    def test_xirr_basic(self):
        self.assertEqual(round(xirr([-100, 30, 30, 30, 30], [43571, 43721, 43871, 44021, 44171], 0), 7), 0.1981947)
        self.assertEqual(round(xirr([-130, 30, 30, 30, 30], [43571, 43721, 43871, 44021, 44171], 0), 7), -0.0743828)
        self.assertIsInstance(xirr(
            [-0.01, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1312.46, -565.86, -711.87, -1410.09, -1495.17, -2943.51, -1450.77,
             -783.22, -112.72, -137.33, 428.64, -1340.84, -256.75],
            [43646, 43830, 44012, 44196, 44377, 44561, 44742, 44926, 45107, 45291, 45473, 45657, 45838, 46022, 46203,
             46387, 46568, 46752, 46934, 47118, 47299, 47483, 47664, 47848, 48029],
            0),
                              ExcelError)  # under this example, Excel would actually return a wrong value


class Test_Npv(unittest.TestCase):
    def test_npv_errors(self):
        self.assertIsInstance(npv(0.06, [1, 2, ExcelError('#NUM')]), ExcelError)
        self.assertIsInstance(npv(ExcelError('#NUM'), [1, 2, 3]), ExcelError)


    def test_npv_basic(self):
        self.assertEqual(round(npv(0.06, [1, 2, 3]), 7), 5.2422470)
        self.assertEqual(round(npv(0.06, 1, 2, 3), 7), 5.2422470)
        self.assertEqual(round(npv(0.06, 1), 7), 0.9433962)


class Test_Offset(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_offset_height_not_integer(self):
        with self.assertRaises(ExcelError):
            offset('Sheet1!A1', 'e', 2)

    @unittest.skip('This test fails.')
    def test_offset_height_is_zero(self):
        with self.assertRaises(ExcelError):
            offset('Sheet1!A1', 1, 2, 0, 1)

    @unittest.skip('This test fails.')
    def test_offset_only_height(self):
        with self.assertRaises(ExcelError):
            offset('Sheet1!A1', 1)

    @unittest.skip('This test fails.')
    def test_offset_out_of_bounds(self):
        with self.assertRaises(ExcelError):
            offset('Sheet1!A1', 1, -2)

    def test_offset_regular(self):
        self.assertEqual(offset('A1:B2', 1, 2), 'C2')

    def test_offset_with_sheet(self):
        self.assertEqual(offset('Sheet1!A1:Sheet1!B2', 1, 2), 'Sheet1!C2')

    def test_offset_rectangular(self):
        self.assertEqual(offset('Sheet1!A1:B2', 1, 2, 2, 3), 'Sheet1!C2:E3')


class Test_Lookup(unittest.TestCase):
    def test_lookup_with_result_range(self):
        range1 = Range('A1:A3', [1, 2, 3])
        range2 = Range('B1:B3', ['blue', 'orange', 'green'])

        self.assertEqual(lookup(2, range1, range2), 'orange')

    def test_lookup_find_closest_inferior(self):
        range = Range('A1:A3', [1, 2, 3])
        self.assertEqual(lookup(2.5, range), 2)

    def test_lookup_basic(self):
        range = Range('A1:A3', [1, 2, 3])
        self.assertEqual(lookup(2, range), 2)


class Test_Average(unittest.TestCase):
    def test_average(self):
        range = Range('A1:A3', [2, 4, 6])
        value = 8
        self.assertEqual(average(range, value), 5)


class Test_Min(unittest.TestCase):
    def test_min_range_and_value(self):
        range1 = Range('A1:A3', [1, 23, 3])
        range2 = Range('A1:A3', [1, 23, "A"])
        value = 20
        self.assertEqual(xmin(range1, value), 1)
        self.assertEqual(xmin(range2, value), 1)


class Test_Max(unittest.TestCase):
    def test_max_range_and_value(self):
        range = Range('A1:A3', [1, 23, 3])
        value = 20
        self.assertEqual(xmax(range, value), 23)


class Test_Rows(unittest.TestCase):
    def test_rows(self):
        range = Range('A1:A3', [1, 23, 3])
        self.assertEqual(rows(range), 3)


class Test_Columns(unittest.TestCase):
    def test_rows(self):
        range = Range('A1:C1', [1, 23, 3])
        self.assertEqual(columns(range), 3)


class Test_Sum(unittest.TestCase):
    def test_sum_ignores_non_numeric(self):
        range = Range('A1:A3', [1, 'e', 3])
        self.assertEqual(xsum(range), 4)

    def test_sum_returns_zero_when_no_numeric(self):
        range = Range('A1:A3', ['ER', 'Er', 're'])
        value = 'ererr'
        self.assertEqual(xsum(range, value), 0)

    def test_sum_excludes_booleans_from_nested_ranges(self):
        range = Range('A1:A3', [True, 2, 1])
        value = True
        self.assertEqual(xsum(range, value), 4)

    def test_sum_range_and_value(self):
        range = Range('A1:A3', [1, 2, 3])
        value = 20
        self.assertEqual(xsum(range, value), 26)


class Test_Iferror(unittest.TestCase): # Need rewriting
    @unittest.skip('This test fails.')
    def test_when_error(self):
        self.assertEqual(iferror('3 / 0', 4), 4)

    @unittest.skip('This test fails.')
    def test_when_no_error(self):
        self.assertEqual(iferror('3 * 10', 4), 30)


class Test_Index(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_index_not_range(self):
        with self.assertRaises(ExcelError):
            index([0, 0.2, 0.15], 2, 4)

    @unittest.skip('This test fails.')
    def test_index_row_not_number(self):
        with self.assertRaises(ExcelError):
            index(resolve_range('A1:A3'), 'e', 1)

    @unittest.skip('This test fails.')
    def test_index_col_not_number(self):
        with self.assertRaises(ExcelError):
            index(resolve_range('A1:D3'), 1, 'e')

    @unittest.skip('This test fails.')
    def test_index_dim_zero(self):
        with self.assertRaises(ExcelError):
            index(resolve_range('A1:D3'), 0, 0)

    def test_index_1_dim_2_coords(self):
        self.assertEqual(index(resolve_range('A1:A3'), 3, 1), 'A3')

    @unittest.skip('This test fails.')
    def test_index_1_dim_out_of_range(self):
        with self.assertRaises(ExcelError):
            index(resolve_range('A1:A3'), 4)

    def test_index_1_dim(self):
        self.assertEqual(index(resolve_range('A1:A3'), 3), 'A3')
        self.assertEqual(index(resolve_range('A1:C1'), 3), 'C1')

    @unittest.skip('This test fails.')
    def test_index_2_dim_1_coord(self):
        with self.assertRaises(ExcelError):
            index(resolve_range('D1:F2'), 2)

    @unittest.skip('This test fails.')
    def test_index_2_dim_out_of_range(self):
        with self.assertRaises(ExcelError):
            index(resolve_range('D1:F2'), 2, 6)

    def test_index_2_dim_row_0(self):
        self.assertEqual(index(resolve_range('D1:F2'), 0, 3), ['F1', 'F2'])

    def test_index_2_dim_col_0(self):
        self.assertEqual(index(resolve_range('D1:F2'), 2, 0), ['D2', 'E2', 'F2'])

    @unittest.skip('This test fails.')
    def test_index_2_dim_col_0_ref_not_found(self):
        # range = Range(['D1', 'E1', 'F1', 'D2', 'E2', 'F2'], [1, 2, 3, 4, 5, 6])

        with self.assertRaises(ExcelError):
            index(resolve_range('D1:F2'), 2, 0)

    def test_index_2_dim(self):
        self.assertEqual(index(resolve_range('D1:F2'), 2, 2), 'E2')


class Test_SumProduct(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_ranges_with_different_sizes(self):
        range1 = Range('A1:A3', [1, 10, 3])
        range2 = Range('B1:B4', [3, 3, 1, 2])

        with self.assertRaises(ExcelError):
            sumproduct(range1, range2)

    def test_regular(self):
        range1 = Range('A1:A3', [1, 10, 3])
        range2 = Range('B1:B3', [3, 3, 1])

        self.assertEqual(sumproduct(range1, range2), 36)


class Test_SumIf(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_range_is_a_list(self):
        with self.assertRaises(ExcelError):
            sumif(12, 12)

    @unittest.skip('This test fails.')
    def test_sum_range_is_a_list(self):
        with self.assertRaises(ExcelError):
            sumif(12, 12, 12)

    def test_criteria_is_not_number_string_boolean(self):
        range1 = Range('A1:A3', [1, 2, 3])
        range2 = Range('A1:A2', [1, 2])

        self.assertEqual(sumif(range1, range2), 0)

    def test_regular_with_number_criteria(self):
        range = Range('A1:A5', [1, 1, 2, 2, 2])

        self.assertEqual(sumif(range, 2), 6)

    def test_regular_with_string_criteria(self):
        range = Range('A1:A6', [1, 2, 3, 4, 5, 6])

        self.assertEqual(sumif(range, ">=3"), 18)

    def test_sum_range(self):
        range1 = Range('A1:A5', [1, 2, 3, 4, 5])
        range2 = Range('A1:A5', [100, 123, 12, 23, 633])

        self.assertEqual(sumif(range1, ">=3", range2), 668)

    def test_sum_range_with_more_indexes(self):
        range1 = Range('A1:A5', [1, 2, 3, 4, 5])
        range2 = Range('A1:A6', [100, 123, 12, 23, 633, 1])

        self.assertEqual(sumif(range1, ">=3", range2), 668)

    def test_sum_range_with_less_indexes(self):
        range1 = Range('A1:A5', [1, 2, 3, 4, 5])
        range2 = Range('A1:A4', [100, 123, 12, 23])

        self.assertEqual(sumif(range1, ">=3", range2), 35)


class Test_SumIfs(unittest.TestCase):

    def test_criteria_numeric(self):
        sum_range = Range('A1:A3', [1, 2, 3])
        criteria_range = Range('B1:B3', [1, 2, 3])

        self.assertEqual(sumifs(sum_range, criteria_range, '<2'), 1)
        self.assertEqual(sumifs(sum_range, criteria_range, '<=2'), 3)
        self.assertEqual(sumifs(sum_range, criteria_range, '>2'), 3)

    def test_criteria_string(self):
        sum_range = Range('A1:A3', [1, 2, 3])
        criteria_range = Range('B1:B3', ['A', 'B', 'C'])

        self.assertEqual(sumifs(sum_range, criteria_range, '=A'), 1)
        self.assertEqual(sumifs(sum_range, criteria_range, '=B'), 2)
        self.assertEqual(sumifs(sum_range, criteria_range, 'C'), 3)

    def test__multiple_criteria(self):
        sum_range = Range('A1:A3', [1, 2, 3])
        criteria_range1 = Range('B1:B3', ['A', 'B', 'B'])
        criteria_range2 = Range('B1:B3', [1, 2, 3])

        self.assertEqual(sumifs(sum_range, criteria_range1, '=A', criteria_range2, '<2'), 1)
        self.assertEqual(sumifs(sum_range, criteria_range1, '=B', criteria_range2, '>1'), 5)
        self.assertEqual(sumifs(sum_range, criteria_range1, 'B', criteria_range2, '<3'), 2)


class Test_IsNa(unittest.TestCase):
    # This function might need more solid testing

    @unittest.skip('This test fails.')
    def test_isNa_false(self):
        self.assertFalse(isNa('2 + 1'))

    @unittest.skip('This test fails.')
    def test_isNa_true(self):
        self.assertTrue(isNa('x + 1'))


class Test_Yearfrac(unittest.TestCase):
    def test_start_date_must_be_number(self):
        self.assertEqual(type(yearfrac('not a number', 1)), ExcelError)

    def test_end_date_must_be_number(self):
        self.assertEqual(type(yearfrac(1, 'not a number')), ExcelError)

    def test_start_date_must_be_positive(self):
        self.assertEqual(type(yearfrac(-1, 0)), ExcelError)

    def test_end_date_must_be_positive(self):
        self.assertEqual(type(yearfrac(0, -1)), ExcelError)

    @unittest.skip('This test fails.')
    def test_basis_must_be_between_0_and_4(self):
        with self.assertRaises(ExcelError):
            yearfrac(1, 2, 5)

    def test_yearfrac_basis_0(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20)), 7.30277777777778)

    def test_yearfrac_basis_1(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 1), 7.299110198)

    def test_yearfrac_basis_2(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 2), 7.405555556)

    def test_yearfrac_basis_3(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 3), 7.304109589)

    def test_yearfrac_basis_4(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 4), 7.302777778)

    def test_yearfrac_inverted(self):
        self.assertAlmostEqual(yearfrac(date(2015, 4, 20), date(2008, 1, 1)), yearfrac(date(2008, 1, 1), date(2015, 4, 20)))


class Test_Date(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_year_must_be_integer(self):
        with self.assertRaises(ExcelError):
            date('2016', 1, 1)

    @unittest.skip('This test fails.')
    def test_month_must_be_integer(self):
        with self.assertRaises(ExcelError):
            date(2016, '1', 1)

    @unittest.skip('This test fails.')
    def test_day_must_be_integer(self):
        with self.assertRaises(ExcelError):
            date(2016, 1, '1')

    @unittest.skip('This test fails.')
    def test_year_must_be_positive(self):
        with self.assertRaises(ExcelError):
            date(-1, 1, 1)

    @unittest.skip('This test fails.')
    def test_year_must_have_less_than_10000(self):
        with self.assertRaises(ExcelError):
            date(10000, 1, 1)

    @unittest.skip('This test fails.')
    def test_result_must_be_positive(self):
        with self.assertRaises(ArithmeticError):
            date(1900, 1, -1)

    def test_not_stricly_positive_month_substracts(self):
        self.assertEqual(date(2009, -1, 1), date(2008, 11, 1))

    def test_not_stricly_positive_day_substracts(self):
        self.assertEqual(date(2009, 1, -1), date(2008, 12, 30))

    def test_month_superior_to_12_change_year(self):
        self.assertEqual(date(2009, 14, 1), date(2010, 2, 1))

    def test_day_superior_to_365_change_year(self):
        self.assertEqual(date(2009, 1, 400), date(2010, 2, 4))

    def test_year_for_29_feb(self):
        self.assertEqual(date(2008, 2, 29), 39507)

    def test_year_regular(self):
        self.assertEqual(date(2008, 11, 3), 39755)


class Test_Mid(unittest.TestCase):
    def test_start_num_must_be_integer(self):
        self.assertIsInstance(mid('Romain', 1.1, 2), ExcelError)  # error is not raised but returned

    def test_num_chars_must_be_integer(self):
        self.assertIsInstance(mid('Romain', 1, 2.1), ExcelError)  # error is not raised but returned

    def test_start_num_must_be_superior_or_equal_to_1(self):
        self.assertIsInstance(mid('Romain', 0, 3), ExcelError)  # error is not raised but returned

    def test_num_chars_must_be_positive(self):
        self.assertIsInstance(mid('Romain', 1, -1), ExcelError)  # error is not raised but returned

    def test_mid(self):
        self.assertEqual(mid('Romain', 3, 4), 'main')
        self.assertEqual(mid('Romain', 1, 2), 'Ro')
        self.assertEqual(mid('Romain', 3, 6), 'main')


class Test_Round(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_nb_must_be_number(self):
        with self.assertRaises(ExcelError):
            round('er', 1)

    @unittest.skip('This test fails.')
    def test_nb_digits_must_be_number(self):
        with self.assertRaises(ExcelError):
            round(2.323, 'ze')

    def test_positive_number_of_digits(self):
        self.assertEqual(xround(2.675, 2), 2.68)

    def test_negative_number_of_digits(self):
        self.assertEqual(xround(2352.67, -2), 2400)


class Test_Count(unittest.TestCase):
    def test_without_nested_booleans(self):
        range = Range('A1:A3', [1, 2, 'e'])

        self.assertEqual(count(range, True, 'r'), 3)

    def test_with_nested_booleans(self):
        range = Range('A1:A3', [1, True, 'e'])

        self.assertEqual(count(range, True, 'r'), 2)

    def test_with_text_representations(self):
        range = Range('A1:A3', [1, '2.2', 'e'])

        self.assertEqual(count(range, True, '20'), 4)


class Test_Countif(unittest.TestCase):
    def setUp(self):
        pass

    @unittest.skip('This test fails.')
    def test_argument_validity(self):
        range = Range('A1:A2', ['e', 1])

        with self.assertRaises(ExcelError):
            countif(range, '>=1')

    def test_countif_strictly_superior(self):
        range = Range('A1:A4', [7, 25, 13, 25])

        self.assertEqual(countif(range, '>10'), 3)

    def test_countif_strictly_inferior(self):
        range = Range('A1:A4', [7, 25, 13, 25])

        self.assertEqual(countif(range, '<10'), 1)

    def test_countif_superior(self):
        range = Range('A1:A4', [7, 25, 13, 25])

        self.assertEqual(countif(range, '>=10'), 3)

    def test_countif_inferior(self):
        range = Range('A1:A4', [7, 25, 10, 25])

        self.assertEqual(countif(range, '<=10'), 2)

    def test_countif_different(self):
        range = Range('A1:A4', [7, 25, 10, 25])

        self.assertEqual(countif(range, '<>10'), 3)

    def test_countif_with_string_equality(self):
        range = Range('A1:A4', [7, 'e', 13, 'e'])

        self.assertEqual(countif(range, 'e'), 2)

    def test_countif_regular(self):
        range = Range('A1:A4', [7, 25, 13, 25])

        self.assertEqual(countif(range, 25), 2)


class Test_Countifs(unittest.TestCase): # more tests might be welcomed
    def setUp(self):
        pass

    @unittest.skip('This test fails.')
    def test_countifs_not_associated(self):  # ASSOCIATION IS NOT TESTED IN COUNTIFS BUT IT SHOULD BE
        range1 = Range('A1:A4', [7, 25, 13, 25])
        range2 = Range('B2:B5', [100, 102, 201, 20])
        range3 = Range('C3:C7', [100, 102, 201, 20])

        with self.assertRaises(ExcelError):
            countifs(range1, 25, range2, ">100", range3, "<=100")

    def test_countifs_regular(self):
        range1 = Range('A1:A4', [7, 25, 13, 25])
        range2 = Range('B1:B4', [100, 102, 201, 20])

        self.assertEqual(countifs(range1, 25, range2, ">100"), 1)


class Test_Mod(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_first_argument_validity(self):
        with self.assertRaises(ExcelError):
            mod(2.2, 1)

    @unittest.skip('This test fails.')
    def test_second_argument_validity(self):
        with self.assertRaises(ExcelError):
            mod(2, 1.1)

    def test_output_value(self):
        self.assertEqual(mod(10, 4), 2)


class Test_Match(unittest.TestCase):
    def test_numeric_in_ascending_mode(self):
        range = Range('A1:A3', [1, 3.3, 5])
        # Closest inferior value is found
        self.assertEqual(match(5, range), 3)

    @unittest.skip('This test fails.')
    def test_numeric_in_ascending_mode_with_descending_array(self):
        range = Range('A1:A4', [10, 9.1, 6.23, 1])
        # Not ascending arrays raise exception
        with self.assertRaises(ExcelError):
            match(3, range)

    @unittest.skip('This test fails.')
    def test_numeric_in_ascending_mode_with_any_array(self):
        range = Range('A1:A4', [10, 3.3, 5, 2])
        # Not ascending arrays raise exception
        with self.assertRaises(ExcelError):
            match(3, range)

    def test_numeric_in_exact_mode(self):
        range = Range('A1:A3', [10, 3.3, 5.0])
        # Value is found
        self.assertEqual(match(5, range, 0), 3)

    @unittest.skip('This test fails.')
    def test_numeric_in_exact_mode_not_found(self):
        range = Range('A1:A4', [10, 3.3, 5, 2])
        # Value not found raises ExcelError
        with self.assertRaises(ExcelError):
            match(3, range, 0)

    def test_numeric_in_descending_mode(self):
        range = Range('A1:A3', [10, 9.1, 6.2])
        # Closest superior value is found
        self.assertEqual(match(8, range, -1), 2)

    @unittest.skip('This test fails.')
    def test_numeric_in_descending_mode_with_ascending_array(self):
        range = Range('A1:A4', [1, 3.3, 5, 6])
        # Non descending arrays raise exception
        with self.assertRaises(ExcelError):
            match(3, range, -1)

    @unittest.skip('This test fails.')
    def test_numeric_in_descending_mode_with_any_array(self):
        range = Range('A1:A4', [10, 3.3, 5, 2])
        # Non descending arrays raise exception
        with self.assertRaises(ExcelError):
            match(3, range, -1)

    def test_string_in_ascending_mode(self):
        range = Range('A1:A3', ['a', 'AAB', 'rars'])
        # Closest inferior value is found
        self.assertEqual(match('rars', range), 3)

    @unittest.skip('This test fails.')
    def test_string_in_ascending_mode_with_descending_array(self):
        range = Range('A1:A3', ['rars', 'aab', 'a'])
        # Not ascending arrays raise exception
        with self.assertRaises(ExcelError):
            match(3, range)

    @unittest.skip('This test fails.')
    def test_string_in_ascending_mode_with_any_array(self):
        range = Range('A1:A3', ['aab', 'a', 'rars'])

        with self.assertRaises(ExcelError):
            match(3, range)

    def test_string_in_exact_mode(self):
        range = Range('A1:A3', ['aab', 'a', 'rars'])
        # Value is found
        self.assertEqual(match('a', range, 0), 2)

    def test_mixed_string_floats_in_exact_mode(self):
        range = Range('A1:A4', ['aab', '3.0', 'rars', 3.3])
        # Value is found
        self.assertEqual(match('aab', range, 0), 1)
        self.assertEqual(match('3.0', range, 0), 2)
        self.assertEqual(match(3, range, 0), 2)
        self.assertEqual(match(3.0, range, 0), 2)
        self.assertEqual(match('rars', range, 0), 3)
        self.assertEqual(match('3.3', range, 0), 4)
        self.assertEqual(match(3.3, range, 0), 4)

    @unittest.skip('This test fails.')
    def test_string_in_exact_mode_not_found(self):
        range = Range('A1:A3', ['aab', 'a', 'rars'])
        # Value not found raises ExcelError
        with self.assertRaises(ExcelError):
            match('b', range, 0)

    def test_string_in_descending_mode(self):
        range = Range('A1:A3', ['c', 'b', 'a'])
        # Closest superior value is found
        self.assertEqual(match('a', range, -1), 3)

    @unittest.skip('This test fails.')
    def test_string_in_descending_mode_with_ascending_array(self):
        range = Range('A1:A3', ['a', 'aab', 'rars'])
        # Non descending arrays raise exception
        with self.assertRaises(ExcelError):
            match('a', range, -1)

    @unittest.skip('This test fails.')
    def test_string_in_descending_mode_with_any_array(self):
        ange = Range('A1:A3', ['aab', 'a', 'rars'])
        # Non descending arrays raise exception
        with self.assertRaises(ExcelError):
            match('a', ['aab', 'a', 'rars'], -1)

    def test_boolean_in_ascending_mode(self):
        range = Range('A1:A3', [False, False, True])
        # Closest inferior value is found
        self.assertEqual(match(True, range), 3)

    @unittest.skip('This test fails.')
    def test_boolean_in_ascending_mode_with_descending_array(self):
        range = Range('A1:A3', [True, False, False])
        # Not ascending arrays raise exception
        with self.assertRaises(ExcelError):
            match(False, range)

    @unittest.skip('This test fails.')
    def test_boolean_in_ascending_mode_with_any_array(self):
        range = Range('A1:A3', [False, True, False])
        # Not ascending arrays raise exception
        with self.assertRaises(ExcelError):
            match(True, range)

    def test_boolean_in_exact_mode(self):
        range = Range('A1:A3', [True, False, False])
        # Value is found
        self.assertEqual(match(False, range, 0), 2)

    @unittest.skip('This test fails.')
    def test_boolean_in_exact_mode_not_found(self):
        range = Range('A1:A3', [True, True, True])
        # Value not found raises ExcelError
        with self.assertRaises(ExcelError):
            match(False, range, 0)

    def test_boolean_in_descending_mode(self):
        range = Range('A1:A3', [True, False, False])
        # Closest superior value is found
        self.assertEqual(match(False, range, -1), 3)

    @unittest.skip('This test fails.')
    def test_boolean_in_descending_mode_with_ascending_array(self):
        range = Range('A1:A3', [False, False, True])
        # Non descending arrays raise exception
        with self.assertRaises(ExcelError):
            match(False, range, -1)

    @unittest.skip('This test fails.')
    def test_boolean_in_descending_mode_with_any_array(self):
        range = Range('A1:A3', [False, True, False])
        with self.assertRaises(ExcelError):
            match(True, range, -1)


# https://support.office.com/en-ie/article/power-function-d3f2908b-56f4-4c3f-895a-07fb519c362a
class Test_Power(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_first_argument_validity(self):
        with self.assertRaises(ExcelError):
            power(-1, 2)

    @unittest.skip('This test fails.')
    def test_second_argument_validity(self):
        with self.assertRaises(ExcelError):
            power(1, 0)

    def test_integers(self):
        self.assertEqual(power(5, 2), 25)

    def test_floats(self):
        self.assertEqual(power(98.6, 3.2), 2401077.2220695773)


# https://support.office.com/en-ie/article/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdfa
class Test_Sqrt(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_first_argument_validity(self):
        with self.assertRaises(ExcelError):
            sqrt(-16)

    def test_positive_integers(self):
        self.assertEqual(sqrt(16), 4)


class Test_Sqrt(unittest.TestCase):
    @unittest.skip('This test fails.')
    def test_first_argument_validity(self):
        with self.assertRaises(ExcelError):
            sqrt(-16)

    def test_positive_integers(self):
        self.assertEqual(sqrt(16), 4)

    def test_float(self):
        self.assertEqual(sqrt(.25), .5)


class Test_Today(unittest.TestCase):

    EXCEL_EPOCH = datetime.datetime.strptime("1900-01-01", '%Y-%m-%d').date()
    reference_date = datetime.datetime.today().date()
    days_since_epoch = reference_date - EXCEL_EPOCH
    todays_ordinal = days_since_epoch.days + 2

    def test_positive_integers(self):
        self.assertEqual(today(), self.todays_ordinal)


class Test_Concatenate(unittest.TestCase):

    @unittest.skip('This test fails.')
    def test_first_argument_validity(self):
        with self.assertRaises(ExcelError):
            concatenate("Hello ", 2, [' World!'])

    def test_concatenate(self):
        self.assertEqual(concatenate("Hello", " ", "World!"), "Hello World!")


class Test_Year(unittest.TestCase):

    def test_results(self):
        self.assertEqual(year(43566), 2019)  # 11/04/2019
        self.assertEqual(year(43831), 2020)  # 01/01/2020
        self.assertEqual(year(36525), 1999)  # 31/12/1999


class Test_Month(unittest.TestCase):

    def test_results(self):
        self.assertEqual(month(43566), 4)  # 11/04/2019
        self.assertEqual(month(43831), 1)  # 01/01/2020
        self.assertEqual(month(36525), 12)  # 31/12/1999


class Test_Eomonth(unittest.TestCase):

    def test_results(self):
        self.assertEqual(eomonth(43566, 2), 43646)  # 11/04/2019, add 2 months
        self.assertEqual(eomonth(43566, 2.1), 43646)  # 11/04/2019, add 2 months
        self.assertEqual(eomonth(43566, 2.99), 43646)  # 11/04/2019, add 2 months
        self.assertEqual(eomonth(43831, 5), 44012)  # 01/01/2020, add 5 months
        self.assertEqual(eomonth(36525, 1), 36556)  # 31/12/1999, add 1 month
        self.assertEqual(eomonth(36525, 15), 36981)  # 31/12/1999, add 15 month
        self.assertNotEqual(eomonth(36525, 15), 36980)  # 31/12/1999, add 15 month
