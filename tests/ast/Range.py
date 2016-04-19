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
        
    def test_get_values(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], [1, 2, 3])

    	self.assertEqual(get_values('C1', range1, range2), (1, 1))

    def test_get_values_raises_error(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], [1, 2, 3])

        with self.assertRaises(Exception):
            get_values('C5', range1, range2)

    # ADD
    def test_add_array(self):
        range1 = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        range2 = Range(['B1', 'B2', 'B3'], [1, 2, 3])

        self.assertEqual(Range.add(range1, range2, 'C1'), 2) # 1 + 1 = 2

    def test_add_array_constant(self):
        range = Range(['A1', 'A2', 'A3'], [1, 2, 3])
        constant = 2

        self.assertEqual(Range.add(range, constant, 'C1'), 3) # 1 + 2 = 3

    # def test_add_constant(self):
    #     self.assertEqual(Range.add(Range([1, 2, 3, 4]), 5, 2), 8) # 3 + 5 = 8

    
    # # SUBSTRACT
    # def test_substract_array(self):
    #     self.assertEqual(Range.substract(Range([1, 2, 3, 4]), Range([4, 3, 2, 1]), 2), 1) # 3 - 2 = 1

    # def test_substract_constant(self):
    #     self.assertEqual(Range.substract(Range([1, 2, 3, 4]), 3, 2), 0) # 3 - 3 = 0

    # # MULTIPLY
    # def test_multiply_array(self):
    #     self.assertEqual(Range.multiply(Range([1, 2, 3, 4]), Range([4, 3, 2, 1]), 2), 6) # 3 * 2 = 6

    # def test_multiply_constant(self):
    #     self.assertEqual(Range.multiply(Range([1, 2, 3, 4]), 3, 2), 9) # 3 * 3 = 9 


    # # DIVIDE
    # def test_divide_array(self):
    #     self.assertEqual(Range.divide(Range([1.0, 2.0, 3.0, 4.0]), Range([4.0, 3.0, 2.0, 1.0]), 2), 1.5) # 3 / 2 = 1.5

    # def test_divide_constant(self):
    #     self.assertEqual(Range.divide(Range([1, 2, 3, 4]), 3, 2), 1) # 3 / 3 = 1

    # # Not sure if this is needed:

    # # # OR
    # # def test_OR_array(self):
    # #     self.assertEqual(Range.OR(Range([True, False]), Range([False, False]), 0), False) # True or False is True

    # # def test_OR_constant(self):
    # #     self.assertEqual(Range.OR(Range([True, False]), False, 0), True)  # True or False if True


    # # # AND
    # # def test_AND_array(self):
    # #     self.assertEqual(Range.AND(Range([True, False]), Range([False, False]), 0), False) # True and False is False

    # # def test_AND_constant(self):
    # #     self.assertEqual(Range.AND(Range([True, False]), False, 0), False)  # True and False if False


    # # IS_EQUAL
    # def test_is_equal_array(self):
    #     self.assertEqual(Range.is_equal(Range([1, 2, 3, 4]), Range([4, 3, 2, 1]), 2), False) # 3 == 2 is False

    # def test_is_equal_constant(self):
    #     self.assertEqual(Range.is_equal(Range([1, 2, 3, 4]), 3, 2), True) # 3 == 3 is True 


    # # IS_NOT_EQUAL
    # def test_is_not_equal_array(self):
    #     self.assertEqual(Range.is_not_equal(Range([1, 2, 3, 4]), Range([4, 3, 2, 1]), 2), True) # 3 != 2 is True

    # def test_is_not_equal_constant(self):
    #     self.assertEqual(Range.is_not_equal(Range([1, 2, 3, 4]), 2, 2), True) # 3 != 2 is True


    # # IS_STRICTLY_SUPERIOR
    # def test_is_strictly_superior_array(self):
    #     self.assertEqual(Range.is_strictly_superior(Range([1, 2, 3, 4]), Range([4, 3, 2, 1]), 2), True) # 3 > 2 is True

    # def test_is_strictly_superior_constant(self):
    #     self.assertEqual(Range.is_strictly_superior(Range([1, 2, 3, 4]), 2, 2), True) # 3 > 2 is True


    # # IS_STRICTLY_INFERIOR
    # def test_is_strictly_inferior_array(self):
    #     self.assertEqual(Range.is_strictly_inferior(Range([1, 2, 3, 4]), Range([4, 3, 2, 1]), 2), False) # 3 < 2 is False

    # def test_is_strictly_inferior_constant(self):
    #     self.assertEqual(Range.is_strictly_inferior(Range([1, 2, 3, 4]), 2, 2), False) # 3 < 2 is False


    # # IS_SUPERIOR_OR_EQUAL
    # def test_is_superior_or_equal_array(self):
    #     self.assertEqual(Range.is_superior_or_equal(Range([1, 2, 3, 4]), Range([3, 3, 3, 3]), 2), True) # 3 >= 3 is True

    # def test_is_superior_or_equal_constant(self):
    #     self.assertEqual(Range.is_superior_or_equal(Range([1, 2, 3, 4]), 2, 2), True) # 3 >= 2 is True


    # # IS_INFERIOR_OR_EQUAL
    # def test_is_inferior_or_equal_array(self):
    #     self.assertEqual(Range.is_inferior_or_equal(Range([1, 2, 3, 4]), Range([3, 3, 3, 3]), 2), True) # 3 <= 3 is True

    # def test_is_inferior_or_equal_constant(self):
    #     self.assertEqual(Range.is_inferior_or_equal(Range([1, 2, 3, 4]), 2, 2), False) # 3 <= 2 is False


if __name__ == '__main__':
    unittest.main()