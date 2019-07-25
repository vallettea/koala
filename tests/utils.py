from types import GeneratorType

import unittest

import koala.utils as utils

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

    # def test_get_linest_degree(self):
    #     """
    #     Testing get_linest_degree
    #     """
    #     pass



if __name__ == '__main__':
    unittest.main()
