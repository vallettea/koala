import unittest
import os
import sys

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../..')
sys.path.insert(0, path)

from koala.ast.Range import Range, get_values

class Test_Excel(unittest.TestCase):
    
    def setUp(self):
        pass
        
    def test_define_Range_with_different_input_sizes(self):
        with self.assertRaises(ValueError):
            Range(['A1', 'A3'], [1, 2, 3])

    def test_get_values(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], [1, 2, 3])

    	self.assertEqual(get_values('C1', range1, range2), (1, 1))

    def test_get_values_raises_error(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], [1, 2, 3])

        with self.assertRaises(Exception):
            get_values('C5', range1, range2)

    def test_is_associated(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], [1, 2, 3])

        self.assertEqual(range1.is_associated(range2), True)

    def test_is_associated_horizontal(self):
        range1 = Range(['A1', 'B1', 'C1'], [1, 2, 3])
        range2 = Range(['A2', 'B2', 'C2'], [1, 2, 3])

        self.assertEqual(range1.is_associated(range2), True)

    def test_is_not_associated(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B2', 'B3', 'B4'], [1, 2, 3])

        self.assertEqual(range1.is_associated(range2), False)

    # ADD
    def test_add_array_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], [1, 2, 3])

        self.assertEqual(Range.add_one(range1, range2, 'C1'), 2) # 1 + 1 = 2

    def test_add_array_one_constant(self):
        range = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        constant = 2

        self.assertEqual(Range.add_one(range, constant, 'C1'), 3) # 1 + 2 = 3

    def test_add_all(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 10, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.add_all(range1, range2, 'C2').values(), [4, 13, 4])

    # SUBSTRACT
    def test_substract_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 10, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.substract_one(range1, range2, 'C2'), 7) # 10 - 3 = 7
    
    # MULTIPLY
    def test_multiply_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 10, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.multiply_one(range1, range2, 'C2'), 30) # 10 * 3 = 30

    # DIVIDE
    def test_divide_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 30, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.divide_one(range1, range2, 'C2'), 10) # 30 / 3 = 10

    # IS_EQUAL
    def test_is_equal_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 30, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.is_equal_one(range1, range2, 'C2'), False) # 30 == 3 is False

    # IS_EQUAL
    def test_is_not_equal_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 30, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.is_not_equal_one(range1, range2, 'C2'), True) # 30 != 3 is True

    # IS_STRICTLY_SUPERIOR
    def test_is_strictly_superior_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 30, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.is_strictly_superior_one(range1, range2, 'C2'), True) # 30 > 3 is True

    # IS_STRICTLY_INFERIOR
    def test_is_strictly_inferior_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 30, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 1])

        self.assertEqual(Range.is_strictly_inferior_one(range1, range2, 'C2'), False) # 30 < 3 is False

    # IS_SUPERIOR_OR_EQUAL
    def test_is_superior_or_equal_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 30, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 3])

        self.assertEqual(Range.is_superior_or_equal_one(range1, range2, 'C3'), True) # 3 >= 3 is True

    # IS_INFERIOR_OR_EQUAL
    def test_is_inferior_or_equal_one(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 30, 3])
        range2 = Range(['B1', 'B2', 'B3'], [3, 3, 3])

        self.assertEqual(Range.is_inferior_or_equal_one(range1, range2, 'C1'), True) # 1 <= 3 is False


if __name__ == '__main__':
    unittest.main()