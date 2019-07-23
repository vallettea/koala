import unittest

import koala.utils as utils

class TestUtil(unittest.TestCase):
    def test_col2num_A(self):
        """
        Test that it can turn 'A' into 1
        """
        data = 'A'
        result = utils.col2num(data)
        self.assertEqual(result, 1)

    def test_num2col_1(self):
        """
        Test that it can turn 1 into 'A'
        """
        data = 1
        result = utils.num2col(data)
        self.assertEqual(result, 'A')

    def test_col2num_BF(self):
        """
        Test that it can turn 'BA' into 53
        """
        data = 'BA'
        result = utils.col2num(data)
        self.assertEqual(result, 53)

    def test_num2col_53(self):
        """
        Test that it can turn 53 into BA
        """
        data = 53
        result = utils.num2col(data)
        self.assertEqual(result, 'BA')

    def test_num2col_0(self):
        """
        Test that it can turn 53 into BA
        """
        data = 0
        with self.assertRaises(Exception) as context:
            utils.num2col(data)

        self.assertTrue('Column ordinal must be larger than 0: 0' in str(context.exception))

    def test_num2col_16385(self):
        """
        Test that it won't go beyond 16384 (which is XFD)
        """
        data = 16385
        with self.assertRaises(Exception) as context:
            utils.num2col(data)

        self.assertTrue('Column ordinal must be less than than 16384: 16385' in str(context.exception))

    def test_col2num_XFE(self):
        """
        Test that it returns an error when columns are beyond 'XFD'
        """
        data = 'XFE'
        with self.assertRaises(Exception) as context:
            utils.col2num(data)

        self.assertTrue('Column ordinal must be left of XFD: XFE' in str(context.exception))

    def test_split_address_Sheet1A1(self):
        """
        Test that it can split address Sheet1!A1
        """
        data = 'Sheet1!A1'
        result = utils.split_address(data)
        self.assertEqual(result, ('Sheet1', 'A', '1'))

    def test_split_address_Sheet1A0(self):
        """
        Test that it can split address Sheet1!A1
        """
        data = 'Sheet1!A0'
        result = utils.split_address(data)
        self.assertEqual(result, ('Sheet1', 'A', '0'))

    def test_split_address_Sheet1XFE1(self):
        """
        Test that it can split address Sheet1!A1
        """
        data = 'Sheet1!XFE1'
        result = utils.split_address(data)
        self.assertEqual(result, ('Sheet1', 'XFE', '1'))

    def test_split_address_Sheet1XFE0(self):
        """
        Test that it can split address Sheet1!A1
        """
        data = 'Sheet1!XFE0'
        result = utils.split_address(data)
        self.assertEqual(result, ('Sheet1', 'XFE', '0'))

if __name__ == '__main__':
    unittest.main()
