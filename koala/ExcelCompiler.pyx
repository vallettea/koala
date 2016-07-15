# cython: profile=True

import os.path
import textwrap
from math import *

import networkx
from networkx.algorithms import number_connected_components

from reader import read_archive, read_named_ranges, read_cells
from excellib import *
from utils import *
from ast import graph_from_seeds, shunting_yard, build_ast
from ExcelError import *
from Cell import Cell
from Range import RangeFactory
from Spreadsheet import Spreadsheet


class ExcelCompiler(object):
    """Class responsible for taking cells and named_range and create a graph
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, file, ignore_sheets = [], ignore_hidden = False, debug = False):
        print "___### Initializing Excel Compiler ###___"

        file_name = os.path.abspath(file)
        # Decompose subfiles structure in zip file
        archive = read_archive(file_name)
        # Parse cells
        self.cells = read_cells(archive, ignore_sheets, ignore_hidden)
        # Parse named_range { name (ExampleName) -> address (Sheet!A1:A10)}
        self.named_ranges = read_named_ranges(archive)
        self.Range = RangeFactory(self.cells)
        self.debug = debug

    def clean_volatile(self):
        print '___### Cleaning volatiles ###___'

        sp = Spreadsheet(networkx.DiGraph(),self.cells, self.named_ranges, debug = self.debug)

        cleaned_cells, cleaned_ranged_names = sp.clean_volatile()
        self.cells = cleaned_cells

        self.named_ranges = cleaned_ranged_names
            
    def gen_graph(self, outputs = None):
        print '___### Generating Graph ###___'

        if outputs is None:
            outputs = set(list(flatten(self.cells.keys())) + self.named_ranges.keys())
        else:
            outputs = set(list(outputs)) # creates a copy
        
        # print 'OUTPUTS', outputs

        seeds = []
        for o in outputs:
            if o in self.named_ranges:
                reference = self.named_ranges[o]

                if is_range(reference):
                    if 'OFFSET' in reference or 'INDEX' in reference:
                        print 'REF', reference
                        start, end = reference.split(':')
                        print 'Start', start, 'End', end

                        def build_code(formula):
                            e = shunting_yard(formula, self.named_ranges, tokenize_range = False)
                            debug = True
                            ast,root = build_ast(e, debug = debug)
                            code = root.emit(ast)

                            print 'Code', code

                        [start_code, end_code] = map(build_code, [start, end])

                        start_end = {
                            "start": start_code,
                            "end": end_code
                        }

                        rng = self.Range(start_end)
                    else:
                        rng = self.Range(reference)

                    # rng = self.Range(reference)
                    for address in rng.addresses: # this is avoid pruning deletion
                        outputs.add(address)
                    virtual_cell = Cell(o, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                    seeds.append(virtual_cell)
                else:
                    # might need to be changed to actual self.cells Cell, not a copy
                    virtual_cell = Cell(o, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True)
                    seeds.append(virtual_cell)
            else:
                if is_range(o):
                    rng = self.Range(o)
                    for address in rng.addresses: # this is avoid pruning deletion
                        outputs.add(address)
                    virtual_cell = Cell(o, None, value = rng, formula = o, is_range = True, is_named_range = True )
                    seeds.append(virtual_cell)
                else:
                    seeds.append(self.cells[o])

        print "Seeds %s cells" % len(seeds)

        # print "%s cells on the todo list" % len(todo)

        cellmap, G = graph_from_seeds(seeds, self)

        print "Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))
        undirected = networkx.Graph(G)
        # print "Number of connected components %s", str(number_connected_components(undirected))

        return Spreadsheet(G, cellmap, self.named_ranges, outputs = outputs, debug = self.debug)




