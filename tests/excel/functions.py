import os
import sys
import unittest

from numpy import array

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../..')
sys.path.insert(0, path)

from koala.ast.excelutils import resolve_range

from koala.ast.excellib import ( 
    xmax,
    xmin,
    xsum,
    average,
    lookup,
    # linest,
    # npv,
    match,
    mod,
    count,
    countif,
    countifs,
    xround,
    mid,
    date,
    yearfrac,
    isNa,
    sumif,
    sumproduct,
    index,
    iferror
)

from koala.ast.Range import Range


class Test_Lookup(unittest.TestCase):
    def setup(self):
        pass

    def test_lookup_with_result_range(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], ['blue', 'orange', 'green'])

        self.assertEqual(lookup(2, range1, range2), 'orange')

    def test_lookup_find_closest_inferior(self):
        range = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        self.assertEqual(lookup(2.5, range), 2)

    def test_lookup_basic(self):
        range = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        self.assertEqual(lookup(2, range), 2)


class Test_Average(unittest.TestCase):
    def setup(self):
        pass

    def test_average(self):
        range = Range(['A1', 'A2', 'A3'], [2, 4, 6])
        value = 8
        self.assertEqual(average(range, value), 5)


class Test_Min(unittest.TestCase):
    def setup(self):
        pass

    def test_min_range_and_value(self):
        range = Range(['A1', 'A2', 'A3'], [1, 23, 3])
        value = 20
        self.assertEqual(xmin(range, value), 1)


class Test_Max(unittest.TestCase):
    def setup(self):
        pass

    def test_max_range_and_value(self):
        range = Range(['A1', 'A2', 'A3'], [1, 23, 3])
        value = 20
        self.assertEqual(xmax(range, value), 23)


class Test_Sum(unittest.TestCase):
    def setup(self):
        pass

    def test_sum_ignores_non_numeric(self):
        range = Range(['A1', 'A2', 'A3'], [1, 'e', 3])
        self.assertEqual(xsum(range), 4)

    def test_sum_returns_zero_when_no_numeric(self):
        range = Range(['A1', 'A2', 'A3'], ['ER', 'Er', 're'])
        value = 'ererr'
        self.assertEqual(xsum(range, value), 0)

    def test_sum_excludes_booleans_from_nested_ranges(self):
        range = Range(['A1', 'A2', 'A3'], [True, 2, 1])
        value = True
        self.assertEqual(xsum(range, value), 4)

    def test_sum_range_and_value(self):
        range = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        value = 20
        self.assertEqual(xsum(range, value), 26)


class Test_Iferror(unittest.TestCase):
    def setup(self):
        pass

    def test_when_error(self):
        self.assertEqual(iferror('3 / 0', 4), 4)

    def test_when_no_error(self):
        self.assertEqual(iferror('3 * 10', 4), 30)
    

class Test_Index(unittest.TestCase):
    def setup(self):
        pass

    # def test_index_not_range(self):
    #     with self.assertRaises(TypeError):
    #         index([0, 0.2, 0.15], 2, 4)

    def test_index_row_not_number(self):
        with self.assertRaises(TypeError):
            index(resolve_range('A1:A3'), 'e', 1)

    def test_index_col_not_number(self):
        with self.assertRaises(TypeError):
            index(resolve_range('A1:D3'), 1, 'e')

    def test_index_dim_zero(self):
        with self.assertRaises(ValueError):
            index(resolve_range('A1:D3'), 0, 0)

    def test_index_1_dim_2_coords(self):
        self.assertEqual(index(resolve_range('A1:A3'), 3, 1), 'A3')

    def test_index_1_dim_out_of_range(self):
        with self.assertRaises(Exception):
            index(resolve_range('A1:A3'), 4)

    def test_index_1_dim(self):
        self.assertEqual(index(resolve_range('A1:A3'), 3), 'A3')

    def test_index_2_dim_1_coord(self):
        with self.assertRaises(ValueError):
            index(resolve_range('D1:F2'), 2)

    def test_index_2_dim_out_of_range(self):
        with self.assertRaises(Exception):
            index(resolve_range('D1:F2'), 2, 6)

    def test_index_2_dim_row_0(self):
        self.assertEqual(index(resolve_range('D1:F2'), 0, 3), ['F1', 'F2'])

    def test_index_2_dim_col_0(self):
        self.assertEqual(index(resolve_range('D1:F2'), 2, 0), ['D2', 'E2', 'F2'])

    # def test_index_2_dim_col_0_ref_not_found(self):
    #     # range = Range(['D1', 'E1', 'F1', 'D2', 'E2', 'F2'], [1, 2, 3, 4, 5, 6])

    #     with self.assertRaises(Exception):
    #         index(resolve_range('D1:F2'), 2, 0)

    def test_index_2_dim(self):
        self.assertEqual(index(resolve_range('D1:F2'), 2, 2), 'E2')
            

class Test_SumProduct(unittest.TestCase):
    def setup(self):
        pass

    def test_ranges_with_different_sizes(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 10, 3])
        range2 = Range(['B1', 'B2', 'B3', 'B4'], [3, 3, 1, 2])

        with self.assertRaises(ValueError):
            sumproduct(range1, range2)

    def test_regular(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 10, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(sumproduct(range1, range2), 36)
            

class Test_SumIf(unittest.TestCase):
    def setup(self):
        pass

    def test_range_is_a_list(self):
        with self.assertRaises(TypeError):
            sumif(12, 12)

    def test_sum_range_is_a_list(self):
        with self.assertRaises(TypeError):
            sumif(12, 12, 12)

    def test_criteria_is_not_number_string_boolean(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['A1', 'A2'], [1, 2])

        self.assertEqual(sumif(range1, range2), 0)

    def test_regular_with_number_criteria(self):
        range = Range(['A1', 'A2', 'A3', 'A4', 'A5'], [1, 1, 2, 2, 2])

        self.assertEqual(sumif(range, 2), 6)

    def test_regular_with_string_criteria(self):
        range = Range(['A1', 'A2', 'A3', 'A4', 'A5', 'A6'], [1, 2, 3, 4, 5, 6])

        self.assertEqual(sumif(range, ">=3"), 18)

    def test_sum_range(self):
        range1 = Range(['A1', 'A2', 'A3', 'A4', 'A5'], [1, 2, 3, 4, 5])
        range2 = Range(['A1', 'A2', 'A3', 'A4', 'A5'], [100, 123, 12, 23, 633])

        self.assertEqual(sumif(range1, ">=3", range2), 668)

    def test_sum_range_with_more_indexes(self):
        range1 = Range(['A1', 'A2', 'A3', 'A4', 'A5'], [1, 2, 3, 4, 5])
        range2 = Range(['A1', 'A2', 'A3', 'A4', 'A5', 'A6'], [100, 123, 12, 23, 633, 1])

        self.assertEqual(sumif(range1, ">=3", range2), 668)

    def test_sum_range_with_less_indexes(self):
        range1 = Range(['A1', 'A2', 'A3', 'A4', 'A5'], [1, 2, 3, 4, 5])
        range2 = Range(['A1', 'A2', 'A3', 'A4'], [100, 123, 12, 23])

        self.assertEqual(sumif(range1, ">=3", range2), 35)
        

class Test_IsNa(unittest.TestCase):
    # This function might need more solid testing
    def setup(self):
        pass

    def test_isNa_false(self):
        self.assertFalse(isNa('2 + 1'))

    def test_isNa_true(self):
        self.assertTrue(isNa('x + 1'))


class Test_Yearfrac(unittest.TestCase):
    def setup(self):
        pass

    def test_start_date_must_be_number(self):
        with self.assertRaises(TypeError):
            yearfrac('not a number', 1)

    def test_end_date_must_be_number(self):
        with self.assertRaises(TypeError):
            yearfrac(1, 'not a number')

    def test_start_date_must_be_positive(self):
        with self.assertRaises(ValueError):
            yearfrac(-1, 0)

    def test_end_date_must_be_positive(self):
        with self.assertRaises(ValueError):
            yearfrac(0, -1)

    def test_basis_must_be_between_0_and_4(self):
        with self.assertRaises(ValueError):
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
    def setup(self):
        pass

    def test_year_must_be_integer(self):
        with self.assertRaises(TypeError):
            date('2016', 1, 1)

    def test_month_must_be_integer(self):
        with self.assertRaises(TypeError):
            date(2016, '1', 1)

    def test_day_must_be_integer(self):
        with self.assertRaises(TypeError):
            date(2016, 1, '1')

    def test_year_must_be_positive(self):
        with self.assertRaises(ValueError):
            date(-1, 1, 1)

    def test_year_must_have_less_than_10000(self):
        with self.assertRaises(ValueError):
            date(10000, 1, 1)

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
    def setUp(self):
        pass

    def test_start_num_must_be_integer(self):
        with self.assertRaises(TypeError):
            mid('Romain', 1.1, 2)

    def test_num_chars_must_be_integer(self):
        with self.assertRaises(TypeError):
            mid('Romain', 1, 2.1)

    def test_start_num_must_be_superior_or_equal_to_1(self):
        with self.assertRaises(ValueError):
            mid('Romain', 0, 3)

    def test_num_chars_must_be_positive(self):
        with self.assertRaises(ValueError):
            mid('Romain', 1, -1)

    def test_mid(self):
        self.assertEqual(mid('Romain', 2, 9), 'main')
        

class Test_Round(unittest.TestCase):
    def setUp(self):
        pass

    def test_nb_must_be_number(self):
        with self.assertRaises(TypeError):
            round('er', 1)

    def test_nb_digits_must_be_number(self):
        with self.assertRaises(TypeError):
            round(2.323, 'ze')

    def test_positive_number_of_digits(self):
        self.assertEqual(xround(2.675, 2), 2.68)

    def test_negative_number_of_digits(self):
        self.assertEqual(xround(2352.67, -2), 2400) 


class Test_Count(unittest.TestCase):
    def setUp(self):
        pass

    def test_without_nested_booleans(self):
        range = Range(['A1', 'A2', 'A3'], [1, 2, 'e'])

        self.assertEqual(count(range, True, 'r'), 3)

    def test_with_nested_booleans(self):
        range = Range(['A1', 'A2', 'A3'], [1, True, 'e'])

        self.assertEqual(count(range, True, 'r'), 2)

    def test_with_text_representations(self):
        range = Range(['A1', 'A2', 'A3'], [1, '2.2', 'e'])

        self.assertEqual(count(range, True, '20'), 4)


class Test_Countif(unittest.TestCase):
    def setUp(self):
        pass

    def test_argument_validity(self):
        range = Range(['A1', 'A2'], ['e', 1])

        with self.assertRaises(TypeError):
            countif(range, '>=1')

    def test_countif_strictly_superior(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 13, 25])

        self.assertEqual(countif(range, '>10'), 3)

    def test_countif_strictly_inferior(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 13, 25])

        self.assertEqual(countif(range, '<10'), 1)

    def test_countif_superior(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 13, 25])

        self.assertEqual(countif(range, '>=10'), 3)

    def test_countif_inferior(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 10, 25])

        self.assertEqual(countif(range, '<=10'), 2)

    def test_countif_different(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 10, 25])

        self.assertEqual(countif(range, '<>10'), 3)

    def test_countif_with_string_equality(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [7, 'e', 13, 'e'])

        self.assertEqual(countif(range, 'e'), 2)

    def test_countif_regular(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 13, 25])

        self.assertEqual(countif(range, 25), 2)


class Test_Countifs(unittest.TestCase): # more tests might be welcomed
    def setUp(self):
        pass

    # def test_countifs_not_associated(self): ASSOCIATION IS NOT TESTED IN COUNTIFS BUT IT SHOULD BE
    #     range1 = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 13, 25])
    #     range2 = Range(['B2', 'B3', 'B4', 'B5'], [100, 102, 201, 20])
    #     range3 = Range(['C3', 'C5', 'C6', 'C7'], [100, 102, 201, 20])

    #     with self.assertRaises(ValueError):
    #         countifs(range1, 25, range2, ">100", range3, "<=100")

    def test_countifs_regular(self):
        range1 = Range(['A1', 'A2', 'A3', 'A4'], [7, 25, 13, 25])
        range2 = Range(['B1', 'B2', 'B3', 'B4'], [100, 102, 201, 20])

        self.assertEqual(countifs(range1, 25, range2, ">100"), 1)

class Test_Mod(unittest.TestCase):
    def setUp(self):
        pass

    def test_first_argument_validity(self):
        with self.assertRaises(TypeError):
            mod(2.2, 1)

    def test_second_argument_validity(self):
        with self.assertRaises(TypeError):
            mod(2, 1.1)

    def test_output_value(self):
        self.assertEqual(mod(10, 4), 2)


class Test_Match(unittest.TestCase):
    def setUp(self):
        pass

    def test_numeric_in_ascending_mode(self):
        range = Range(['A1', 'A2', 'A3'], [1, 3.3, 5])
        # Closest inferior value is found
        self.assertEqual(match(5, range), 3)

    def test_numeric_in_ascending_mode_with_descending_array(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [10, 9.1, 6.23, 1])
        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(3, range)

    def test_numeric_in_ascending_mode_with_any_array(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [10, 3.3, 5, 2])
        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(3, range)

    def test_numeric_in_exact_mode(self):
        range = Range(['A1', 'A2', 'A3'], [10, 3.3, 5.0])
        # Value is found
        self.assertEqual(match(5, range, 0), 3)

    def test_numeric_in_exact_mode_not_found(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [10, 3.3, 5, 2])
        # Value not found raises Exception
        with self.assertRaises(ValueError):
            match(3, range, 0)

    def test_numeric_in_descending_mode(self):
        range = Range(['A1', 'A2', 'A3'], [10, 9.1, 6.2])
        # Closest superior value is found
        self.assertEqual(match(8, range, -1), 2)

    def test_numeric_in_descending_mode_with_ascending_array(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [1, 3.3, 5, 6])
        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match(3, range, -1)

    def test_numeric_in_descending_mode_with_any_array(self):
        range = Range(['A1', 'A2', 'A3', 'A4'], [10, 3.3, 5, 2])
        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match(3, range, -1)

    def test_string_in_ascending_mode(self):
        range = Range(['A1', 'A2', 'A3'], ['a', 'AAB', 'rars'])
        # Closest inferior value is found
        self.assertEqual(match('rars', range), 3)

    def test_string_in_ascending_mode_with_descending_array(self):
        range = Range(['A1', 'A2', 'A3'], ['rars', 'aab', 'a'])
        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(3, range)

    def test_string_in_ascending_mode_with_any_array(self):
        range = Range(['A1', 'A2', 'A3'], ['aab', 'a', 'rars'])

        with self.assertRaises(Exception):
            match(3, range)

    def test_string_in_exact_mode(self):
        range = Range(['A1', 'A2', 'A3'], ['aab', 'a', 'rars'])
        # Value is found
        self.assertEqual(match('a', range, 0), 2)

    def test_string_in_exact_mode_not_found(self):
        range = Range(['A1', 'A2', 'A3'], ['aab', 'a', 'rars'])
        # Value not found raises Exception
        with self.assertRaises(ValueError):
            match('b', range, 0)

    def test_string_in_descending_mode(self):
        range = Range(['A1', 'A2', 'A3'], ['c', 'b', 'a'])
        # Closest superior value is found
        self.assertEqual(match('a', range, -1), 3)

    def test_string_in_descending_mode_with_ascending_array(self):
        range = Range(['A1', 'A2', 'A3'], ['a', 'aab', 'rars'])
        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match('a', range, -1)

    def test_string_in_descending_mode_with_any_array(self):
        ange = Range(['A1', 'A2', 'A3'], ['aab', 'a', 'rars'])
        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match('a', ['aab', 'a', 'rars'], -1)

    def test_boolean_in_ascending_mode(self):
        range = Range(['A1', 'A2', 'A3'], [False, False, True])
        # Closest inferior value is found
        self.assertEqual(match(True, range), 3)

    def test_boolean_in_ascending_mode_with_descending_array(self):
        range = Range(['A1', 'A2', 'A3'], [True, False, False])
        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(False, range)

    def test_boolean_in_ascending_mode_with_any_array(self):
        range = Range(['A1', 'A2', 'A3'], [False, True, False])
        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(True, range)

    def test_boolean_in_exact_mode(self):
        range = Range(['A1', 'A2', 'A3'], [True, False, False])
        # Value is found
        self.assertEqual(match(False, range, 0), 2)

    def test_boolean_in_exact_mode_not_found(self):
        range = Range(['A1', 'A2', 'A3'], [True, True, True])
        # Value not found raises Exception
        with self.assertRaises(ValueError):
            match(False, range, 0)

    def test_boolean_in_descending_mode(self):
        range = Range(['A1', 'A2', 'A3'], [True, False, False])
        # Closest superior value is found
        self.assertEqual(match(False, range, -1), 3)

    def test_boolean_in_descending_mode_with_ascending_array(self):
        range = Range(['A1', 'A2', 'A3'], [False, False, True])
        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match(False, range, -1)

    def test_boolean_in_descending_mode_with_any_array(self):
        range = Range(['A1', 'A2', 'A3'], [False, True, False])  
        with self.assertRaises(Exception):
            match(True, range, -1)
 
if __name__ == '__main__':
    unittest.main()
