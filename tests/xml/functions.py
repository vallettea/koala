# from nose.tools import *
import unittest
import numpy

from koala.xml.functions import (
    cell2vec,
    vec2cell,
    translate
)


class Test_SumIf(unittest.TestCase):
    def setup(self):
        pass

    def test_cell2vec(self):
        assert numpy.all([cell2vec('AF13'), numpy.array([13, 32])])

    def test_vec2cell(self):
        assert 'AF13' == vec2cell(numpy.array([13, 32]))

    def test_translate(self):
        assert translate('D1', [2, 7]) == 'K3'

    def test_translate_with_fixed(self):
        assert translate('$D$1', [2, 7]) == '$D$1'

# import pytest

# from koala.xml.functions import ConditionalElement


# @pytest.fixture
# def root():
#     from koala.xml.functions import Element
#     return Element("root")


# @pytest.mark.parametrize("condition", [True, 1, -1])
# def test_simple(root, condition):
#     ConditionalElement(root, "start", condition)
#     assert root.find("start").tag == "start"


# def test_simple_attrib(root):
#     ConditionalElement(root, "start", True, 'val')
#     tag = root.find("start")
#     assert tag.attrib == {'val': '1'}


# def test_dict_attrib(root):
#     ConditionalElement(root, "start", True, {'val':'single'})
#     tag = root.find("start")
#     assert tag.attrib == {'val':'single'}


# @pytest.mark.parametrize("condition", [False, 0, None])
# def test_no_tag(root, condition):
#     ConditionalElement(root, "start", condition)
#     assert root.find("start") is None


# def test_safe_iterator_none():
#     from .. functions import safe_iterator
#     seq = safe_iterator(None)
#     assert seq == []


# @pytest.mark.parametrize("xml, tag",
#                          [
#                              ("<root xmlns='http://ants.builders/ns' />", "root"),
#                              ("<root />", "root"),
#                          ]
#                          )
# def test_localtag(xml, tag):
#     from .. functions import localname
#     from .. functions import fromstring
#     node = fromstring(xml)
#     assert localname(node) == tag





