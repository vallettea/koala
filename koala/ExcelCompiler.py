from __future__ import print_function
# cython: profile=True

import os.path

import networkx

from koala.reader import read_archive, read_named_ranges, read_cells
from koala.utils import *
from koala.ast import graph_from_seeds, shunting_yard, build_ast, prepare_pointer
from koala.Cell import Cell
from koala.Range import RangeFactory
from koala.Spreadsheet import Spreadsheet


class ExcelCompiler(object):
    """Class responsible for taking cells and named_range and create a graph
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, file, ignore_sheets = [], ignore_hidden = False, debug = False):
        # print("___### Initializing Excel Compiler ###___")

        file_name = os.path.abspath(file)
        # Decompose subfiles structure in zip file
        archive = read_archive(file_name)
        # Parse cells
        self.cells = read_cells(archive, ignore_sheets, ignore_hidden)
        # Parse named_range { name (ExampleName) -> address (Sheet!A1:A10)}
        self.named_ranges = read_named_ranges(archive)
        self.Range = RangeFactory(self.cells)
        self.pointers = set()
        self.debug = debug

    def clean_pointer(self):
        sp = Spreadsheet(networkx.DiGraph(),self.cells, self.named_ranges, debug = self.debug)

        cleaned_cells, cleaned_ranged_names = sp.clean_pointer()
        self.cells = cleaned_cells
        self.named_ranges = cleaned_ranged_names
        self.pointers = set()

    def gen_graph(self, outputs = [], inputs = []):
        # print('___### Generating Graph ###___')

        if len(outputs) == 0:
            preseeds = set(list(flatten(self.cells.keys())) + list(self.named_ranges.keys())) # to have unicity
        else:
            preseeds = set(outputs)

        preseeds = list(preseeds) # to be able to modify the list

        seeds = []
        for o in preseeds:
            if o in self.named_ranges:
                reference = self.named_ranges[o]

                if is_range(reference):
                    if 'OFFSET' in reference or 'INDEX' in reference:
                        start_end = prepare_pointer(reference, self.named_ranges)
                        rng = self.Range(start_end)
                        self.pointers.add(o)
                    else:
                        rng = self.Range(reference)

                    for address in rng.addresses: # this is avoid pruning deletion
                        preseeds.append(address)
                    virtual_cell = Cell(o, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                    seeds.append(virtual_cell)
                else:
                    # might need to be changed to actual self.cells Cell, not a copy
                    if 'OFFSET' in reference or 'INDEX' in reference:
                        self.pointers.add(o)

                    value = self.cells[reference].value if reference in self.cells else None
                    virtual_cell = Cell(o, None, value = value, formula = reference, is_range = False, is_named_range = True)
                    seeds.append(virtual_cell)
            else:
                if is_range(o):
                    rng = self.Range(o)
                    for address in rng.addresses: # this is avoid pruning deletion
                        preseeds.append(address)
                    virtual_cell = Cell(o, None, value = rng, formula = o, is_range = True, is_named_range = True )
                    seeds.append(virtual_cell)
                else:
                    seeds.append(self.cells[o])

        seeds = set(seeds)
        # print("Seeds %s cells" % len(seeds))
        outputs = set(preseeds) if len(outputs) > 0 else [] # seeds and outputs are the same when you don't specify outputs

        cellmap, G = graph_from_seeds(seeds, self)

        if len(inputs) != 0: # otherwise, we'll set inputs to cellmap inside Spreadsheet
            inputs = list(set(inputs))

            # add inputs that are outside of calculation chain
            for i in inputs:
                if i not in cellmap:
                    if i in self.named_ranges:
                        reference = self.named_ranges[i]
                        if is_range(reference):

                            rng = self.Range(reference)
                            for address in rng.addresses: # this is avoid pruning deletion
                                inputs.append(address)
                            virtual_cell = Cell(i, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                            cellmap[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain

                        else:
                            # might need to be changed to actual self.cells Cell, not a copy
                            virtual_cell = Cell(i, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True)
                            cellmap[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                    else:
                        if is_range(i):
                            rng = self.Range(i)
                            for address in rng.addresses: # this is avoid pruning deletion
                                inputs.append(address)
                            virtual_cell = Cell(i, None, value = rng, formula = o, is_range = True, is_named_range = True )
                            cellmap[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                        else:
                            cellmap[i] = self.cells[i]
                            G.add_node(self.cells[i]) # edges are not needed here since the input here is not in the calculation chain

            inputs = set(inputs)


        # print("Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap)))

        # undirected = networkx.Graph(G)
        # print "Number of connected components %s", str(number_connected_components(undirected))

        return Spreadsheet(G, cellmap, self.named_ranges, pointers = self.pointers, outputs = outputs, inputs = inputs, debug = self.debug)
