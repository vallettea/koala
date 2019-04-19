# cython: profile=True

from __future__ import absolute_import, division

from koala.CellBase import CellBase
from koala.Range import RangeCore
from koala.utils import *

from openpyxl.compat import unicode


class Cell(CellBase):
    ctr = 0
    __named_range = None

    @classmethod
    def next_id(cls):
        cls.ctr += 1
        return cls.ctr

    def __init__(
            self, address,
            sheet=None, value=None, formula=None,
            is_range=False, is_named_range=False,
            should_eval='normal'):
        super(Cell, self).__init__()

        if not is_named_range:

            # remove $'s
            address = address.replace('$', '')

            sh, c, r = split_address(address)

            # both are empty
            if not sheet and not sh:
                raise Exception(
                    "Sheet name may not be empty for cell address %s" %
                    address)
            # both exist but disagree
            elif sh and sheet and sh != sheet:
                raise Exception(
                    "Sheet name mismatch for cell address %s: %s vs %s" %
                    (address, sheet, sh))
            elif not sh and sheet:
                sh = sheet
            else:
                pass

            # we assume a cell's location can never change
            self.__sheet = sheet.encode('utf-8') if sheet is not None else sheet

            self.__sheet = sh
            self.__col = c
            self.__row = int(r)
            self.__col_idx = col2num(c)

        else:
            self.__named_range = address
            self.__sheet = None
            self.__col = None
            self.__row = None
            self.__col_idx = None

        # `unicode` != `str` in Python2. See `from openpyxl.compat import unicode`
        if type(formula) == str and str != unicode:
            self.__formula = unicode(formula, 'utf-8') if formula else None
        else:
            self.__formula = formula if formula else None

        self.__value = value
        self.python_expression = None
        if (formula is not None) or is_range:
            self.need_update = True
        else:
            self.need_update = False
        self.should_eval = should_eval
        self.__compiled_expression = None
        self.__is_range = is_range

        self.is_range = self.__is_range
        self.is_named_range = self.__named_range is not None

        self.__absolute_address = (
            "%s!%s%s" % (self.__sheet, self.__col, self.__row))
        self.__address = "%s%s" % (self.__col, self.__row)

        # every cell has a unique id
        self.__id = Cell.next_id()

    @property
    def value(self):
        if self.__is_range:
            return self.__value.values
        else:
            return self.__value

    @value.setter
    def value(self, new_value):
        if self.__is_range:
            self.__value.values = new_value
        else:
            self.__value = new_value

    @property
    def range(self):
        if self.__is_range:
            return self.__value
        else:
            return None

    @range.setter
    def range(self, new_range):
        if self.__is_range:
            self.__value = new_range
        else:
            raise Exception(
                'Setting a range as non-range Cell %s value' % self.address())

    @property
    def sheet(self):
        return self.__sheet

    @property
    def row(self):
        return self.__row

    @property
    def col(self):
        return self.__col

    @property
    def formula(self):
        return self.__formula

    @formula.setter
    def formula(self, new_formula):
        # maybe some kind of check is necessary
        self.__formula = new_formula

    @property
    def id(self):
        return self.__id

    @property
    def index(self):
        return self.__index

    @property
    def compiled_expression(self):
        return self.__compiled_expression

    @compiled_expression.setter
    def compiled_expression(self, ce):
        self.__compiled_expression = ce

    # code objects are not serializable
    def __getstate__(self):
        d = dict(self.__dict__)
        f = '__compiled_expression'
        if f in d:
            del d[f]
        return d

    def __setstate__(self, d):
        self.__dict__.update(d)
        self.compile()

    def clean_name(self):
        return self.address().replace('!', '_').replace(' ', '_')

    def address(self, absolute=True):
        if self.is_named_range:
            return self.__named_range
        elif absolute:
            return self.__absolute_address
        else:
            return self.__address

    def address_parts(self):
        return (self.__sheet, self.__col, self.__row, self.__col_idx)

    def compile(self):
        if not self.python_expression:
            return

        try:
            self.__compiled_expression = compile(
                self.python_expression, '<string>', 'eval')
        except Exception as e:
            raise Exception(
                "Failed to compile cell %s with expression %s: %s\nFormula: %s"
                % (self.address(), self.python_expression, e, self.formula))

    def __str__(self):
        return self.address()

    @staticmethod
    def inc_col_address(address, inc):
        sh, col, row = split_address(address)
        return "%s!%s%s" % (sh, num2col(col2num(col) + inc), row)

    @staticmethod
    def inc_row_address(address, inc):
        sh, col, row = split_address(address)
        return "%s!%s%s" % (sh, col, row + inc)

    @staticmethod
    def resolve_cell(excel, address, sheet=None):
        r = excel.get_range(address)
        f = r.Formula if r.Formula.startswith('=') else None
        v = r.Value

        sh, c, r = split_address(address)

        # use the sheet specified in the cell, else the passed sheet
        if sh:
            sheet = sh

        c = Cell(address, sheet, value=v, formula=f)
        return c

    @staticmethod
    def make_cells(excel, range, sheet=None):
        cells = []

        if is_range(range):
            # use the sheet specified in the range, else the passed sheet
            sh, start, end = split_range(range)
            if sh:
                sheet = sh

            ads, numrows, numcols = resolve_range(range)
            # ensure in the same nested format as fs/vs will be
            if numrows == 1:
                ads = [ads]
            elif numcols == 1:
                ads = [[x] for x in ads]

            # get everything in blocks, is faster
            r = excel.get_range(range)
            fs = r.Formula
            vs = r.Value

            for it in (list(zip(*x)) for x in zip(ads, fs, vs)):
                row = []
                for c in it:
                    a = c[0]
                    f = c[1] if c[1] and c[1].startswith('=') else None
                    v = c[2]
                    cl = Cell(a, sheet, value=v, formula=f)
                    row.append(cl)
                cells.append(row)

            # return as vector
            if numrows == 1:
                cells = cells[0]
            elif numcols == 1:
                cells = [x[0] for x in cells]
            else:
                pass
        else:
            c = Cell.resolve_cell(excel, range, sheet=sheet)
            cells.append(c)

            numrows = 1
            numcols = 1

        return (cells, numrows, numcols)

    def asdict(self):
        if self.is_range:
            cell_range = self.range
            value = {
                "cells": cell_range.addresses,
                "values": cell_range.values,
                "nrows": cell_range.nrows,
                "ncols": cell_range.ncols
            }
        else:
            value = self.value

        data = {
            "address": self.address(),
            "formula": self.formula,
            "value": value,
            "python_expression": self.python_expression,
            "is_named_range": self.is_named_range,
            "should_eval": self.should_eval
        }
        return data

    @staticmethod
    def from_dict(d, cellmap=None):
        cell_is_range = type(d["value"]) == dict
        if cell_is_range:
            range = d["value"]
            if len(range["values"]) == 0:
                range["values"] = [None] * len(range["cells"])
            value = RangeCore(
                range["cells"], range["values"],
                nrows=range["nrows"], ncols=range["ncols"], cellmap=cellmap)
        else:
            value = d["value"]
        new_cell = Cell(
            d["address"], None, value=value, formula=d["formula"],
            is_range=cell_is_range,
            is_named_range=d["is_named_range"],
            should_eval=d["should_eval"])
        new_cell.python_expression = d["python_expression"]
        new_cell.compile()

        return new_cell
