#
# source: https://bitbucket.org/openpyxl/openpyxl/src/93604327bce7aac5e8270674579af76d390e09c0/openpyxl/xml/__init__.py?at=default&fileviewer=file-view-default
#_______________________________________________________________________________________________________________________________________________________________

from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl


"""Collection of XML resources compatible across different Python versions"""
import os


def lxml_available():
    try:
        from lxml.etree import LXML_VERSION
        LXML = LXML_VERSION >= (3, 3, 1, 0)
        if not LXML:
            import warnings
            warnings.warn("The installed version of lxml is too old to be used with openpyxl")
            return False  # we have it, but too old
        else:
            return True  # we have it, and recent enough
    except ImportError:
        return False  # we don't even have it


def lxml_env_set():
    return os.environ.get("OPENPYXL_LXML", "True") == "True"


LXML = lxml_available() and lxml_env_set()
