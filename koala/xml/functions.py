from __future__ import absolute_import

"""
XML compatability functions
"""

# Python stdlib imports
import re
import os
import string
from numpy import array
from functools import partial

from lxml.etree import (
    Element,
    ElementTree,
    SubElement,
    fromstring,
    tostring,
    register_namespace,
    iterparse,
    QName,
    xmlfile
)

from xml.etree import (
    cElementTree
)
   

from koala.xml.constants import (
    CHART_NS,
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_DRAWING_NS,
    SHEET_MAIN_NS,
    REL_NS,
    VTYPES_NS,
    COREPROPS_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX
)

# allow LXML interface
_iterparse = iterparse
def safe_iterparse(source, *args, **kw):
    return _iterparse(source)

iterparse = safe_iterparse


register_namespace(DCTERMS_PREFIX, DCTERMS_NS)
register_namespace('dcmitype', 'http://purl.org/dc/dcmitype/')
register_namespace('cp', COREPROPS_NS)
register_namespace('c', CHART_NS)
register_namespace('a', DRAWING_NS)
register_namespace('s', SHEET_MAIN_NS)
register_namespace('r', REL_NS)
register_namespace('vt', VTYPES_NS)
register_namespace('xdr', SHEET_DRAWING_NS)
register_namespace('cdr', CHART_DRAWING_NS)


tostring = partial(tostring, encoding="utf-8")


def safe_iterator(node, tag=None):
    """Return an iterator that is compatible with Python 2.6"""
    if node is None:
        return []
    if hasattr(node, "iter"):
        return node.iter(tag)
    else:
        return node.getiterator(tag)


def ConditionalElement(node, tag, condition, attr=None):
    """
    Utility function for adding nodes if certain criteria are fulfilled
    An optional attribute can be passed in which will always be serialised as '1'
    """
    sub = partial(SubElement, node, tag)
    if bool(condition):
        if isinstance(attr, str):
            elem = sub({attr:'1'})
        elif isinstance(attr, dict):
            elem = sub(attr)
        else:
            elem = sub()
        return elem


NS_REGEX = re.compile("({(?P<namespace>.*)})?(?P<localname>.*)")

def localname(node):
    m = NS_REGEX.match(node.tag)
    return m.group('localname')

def col2num(col): # http://stackoverflow.com/questions/7261936/convert-an-excel-or-spreadsheet-column-letter-to-its-number-in-pythonic-fashion
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def num2col(num): # http://stackoverflow.com/questions/23861680/convert-spreadsheet-number-to-column-letter
    div = num
    string = ""

    while div > 0:
        module = (div - 1) % 26
        string = chr(65 + module) + string
        div = int((div - module) / 26)

    return string


def cell2vec(cell):
    # need to verify match exists

    found = re.search("\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$", cell).group

    row = int(found(2))
    col = int(col2num(found(1)))

    return array([row, col])


def vec2cell(vector):
    # need to verify type(vector) == numpy.ndarray or list

    return num2col(vector[1]) + str(vector[0])


