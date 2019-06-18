from __future__ import absolute_import, print_function
# cython: profile=True

from koala.Range import get_cell_address, parse_cell_address

from koala.ast import *
from koala.reader import read_archive, read_named_ranges, read_cells
# This import equivalent functions defined in Excel.
from koala.excellib import *
from openpyxl.formula.translate import Translator
from koala.serializer import *
from koala.tokenizer import reverse_rpn
from koala.utils import *

import warnings
import os.path
import networkx
from networkx.readwrite import json_graph

from openpyxl.compat import unicode


class Spreadsheet(object):
    def __init__(self, file=None, ignore_sheets=[], ignore_hidden=False, debug=False):
        # print("___### Initializing Excel Compiler ###___")

        if file is None:
            # create empty version of this object
            self.cells = None  # precursor for cellmap: dict that link addresses (str) to Cell objects.
            self.named_ranges = {}
            self.pointers = set()  # set listing the pointers
            self.debug = None  # boolean

            seeds = []
            cellmap, G = graph_from_seeds(seeds, self)
            self.G = G  # DiGraph object that represents the view of the Spreadsheet
            self.cellmap = cellmap  # dict that link addresses (str) to Cell objects.
            self.addr_to_name = None
            self.addr_to_range = None
            self.outputs = None
            self.inputs = None
            self.save_history = None
            self.history = None
            self.count = None
            self.range = RangeFactory(cellmap)
            self.pointer_to_remove = None
            self.pointers_to_reset = set()
            self.reset_buffer = None
            self.fixed_cells = {}
        else:
            # fill in what the ExcelCompiler used to do
            super(Spreadsheet, self).__init__() # generate an empty spreadsheet
            # Decompose subfiles structure in zip file
            if hasattr(file, 'read'):   # file-like object
                archive = read_archive(file)
            else:                       # assume file path
                archive = read_archive(os.path.abspath(file))
            # Parse cells
            self.cells = read_cells(archive, ignore_sheets, ignore_hidden)
            # Parse named_range { name (ExampleName) -> address (Sheet!A1:A10)}
            self.named_ranges = read_named_ranges(archive)
            self.range = RangeFactory(self.cells)
            self.pointers = set()
            self.debug = debug

            # now add the stuff what was originally done by the Spreadsheet
            self.gen_graph()

    def clean_pointer(self):
        spreadsheet = Spreadsheet()
        sp = spreadsheet.build_spreadsheet(networkx.DiGraph(),self.cells, self.named_ranges, debug = self.debug)

        cleaned_cells, cleaned_ranged_names = sp.clean_pointer()
        self.cells = cleaned_cells
        self.named_ranges = cleaned_ranged_names
        self.pointers = set()

    def gen_graph(self, outputs=[], inputs=[]):
        """
        Generate the contents of the Spreadsheet from the read cells in the binary files.
        Specifically this function generates the graph.

        :param outputs: can be used to specify the outputs. All not affected cells are removed from the graph.
        :param inputs: can be used to specify the inputs. All not affected cells are removed from the graph.
        """
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
                        rng = self.range(start_end)
                        self.pointers.add(o)
                    else:
                        rng = self.range(reference)

                    for address in rng.addresses: # this is avoid pruning deletion
                        preseeds.append(address)
                    virtual_cell = Cell(o, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                    seeds.append(virtual_cell)
                else:
                    # might need to be changed to actual cells Cell, not a copy
                    if 'OFFSET' in reference or 'INDEX' in reference:
                        self.pointers.add(o)

                    value = self.cells[reference].value if reference in self.cells else None
                    virtual_cell = Cell(o, None, value = value, formula = reference, is_range = False, is_named_range = True)
                    seeds.append(virtual_cell)
            else:
                if is_range(o):
                    rng = self.range(o)
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

                            rng = self.range(reference)
                            for address in rng.addresses: # this is avoid pruning deletion
                                inputs.append(address)
                            virtual_cell = Cell(i, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                            cellmap[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain

                        else:
                            # might need to be changed to actual cells Cell, not a copy
                            virtual_cell = Cell(i, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True)
                            cellmap[i] = virtual_cell
                            G.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                    else:
                        if is_range(i):
                            rng = self.range(i)
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

        self.build_spreadsheet(G, cellmap, self.named_ranges, pointers = self.pointers, outputs = outputs, inputs = inputs, debug = self.debug)

    def build_spreadsheet(self, G, cellmap, named_ranges, pointers = set(), outputs = set(), inputs = set(), debug = False):
        """
        Writes the elements created by gen_graph to the object

        :param G:
        :param cellmap:
        :param named_ranges:
        :param pointers:
        :param outputs:
        :param inputs:
        :param debug:
        """

        self.G = G
        self.cellmap = cellmap
        self.named_ranges = named_ranges

        addr_to_name = {}
        for name in named_ranges:
            addr_to_name[named_ranges[name]] = name
        self.addr_to_name = addr_to_name

        addr_to_range = {}

        for c in list(self.cellmap.values()):
            if c.is_range and len(list(c.range.keys())) != 0: # could be better, but can't check on Exception types here...
                addr = c.address() if c.is_named_range else c.range.name
                for cell in c.range.addresses:
                    if cell not in addr_to_range:
                        addr_to_range[cell] = [addr]
                    else:
                        addr_to_range[cell].append(addr)

        self.addr_to_range = addr_to_range

        self.outputs = outputs
        self.inputs = inputs
        self.save_history = False
        self.history = dict()
        self.count = 0
        self.pointer_to_remove = ["INDEX", "OFFSET"]
        self.pointers = pointers
        self.pointers_to_reset = pointers
        self.range = RangeFactory(cellmap)
        self.reset_buffer = set()
        self.debug = debug
        self.fixed_cells = {}

        # make sure that all cells that don't have a value defined are updated.
        for cell in self.cellmap.values():
            if cell.value is None and cell.formula is not None:
                cell.needs_update = True


    def activate_history(self):
        self.save_history = True

    def add_cell(self, cell, value = None):
        """
        Depricated, see cell_add().
        """

        if type(cell) != Cell:
            cell = Cell(cell, None, value = value, formula = None, is_range = False, is_named_range = False)

        # previously reset was used to only reset one cell. Capture this behaviour.
        warnings.warn(
            "xxx_cell functions are depricated and replaced by cell_xxx functions. Please use those functions instead. "
            "This behaviour will be removed in a future version.",
            PendingDeprecationWarning
        )
        self.cell_add(cell=cell)

    def cell_add(self, address=None, cell=None, value=None, formula=None):
        """
        Adds a cell to the Spreadsheet. Either the cell argument can be specified, or any combination of the other
        arguments.

        :param address: the address of the cell
        :param cell: a Cell object to add
        :param value: (optional) a new value for the cell. In this case, the first argument cell is processed as
                      an address.
        :param formula:
        """
        if cell is None:
            cell = Cell(address, value=value, formula=formula)

        if address in self.cellmap:
            raise Exception('Cell %s already in cellmap' % address)

        cellmap, G = graph_from_seeds([cell], self)

        self.cellmap = cellmap
        self.G = G

        print("Graph construction updated, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap)))

    def set_formula(self, addr, formula):
        # previously set_formula was used. Capture this behaviour.
        warnings.warn(
            "This function is depricated and will be replaced by cell_set_formula. Please use this function instead. "
            "This behaviour will be removed in a future version.",
            PendingDeprecationWarning
        )
        return self.cell_set_formula(addr, formula)

    def cell_set_formula(self, address, formula):
        """
        Set the formula of a cell.

        :param address: the address of a cell
        :param formula: the new formula
        """
        if address in self.cellmap:
            cell = self.cellmap[address]
        else:
            raise Exception('Cell %s not in cellmap' % address)

        seeds = [cell]

        if cell.is_range:
            for index, c in enumerate(cell.range.cells): # for each cell of the range, translate the formula
                if index == 0:
                    c.formula = formula
                    translator = Translator(unicode('=' +    formula), c.address().split('!')[1]) # the Translator needs a reference without sheet
                else:
                    translated = translator.translate_formula(c.address().split('!')[1]) # the Translator needs a reference without sheet
                    c.formula = translated[1:] # to get rid of the '='

                seeds.append(c)
        else:
            cell.formula = formula

        cellmap, G = graph_from_seeds(seeds, self)

        self.cellmap = cellmap
        self.G = G

        should_eval = self.cellmap[address].should_eval
        self.cellmap[address].should_eval = 'always'
        self.evaluate(address)
        self.cellmap[address].should_eval = should_eval

        print("Graph construction updated, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap)))


    def prune_graph(self, *args):
        print('___### Pruning Graph ###___')

        G = self.G

        # get all the cells impacted by inputs
        dependencies = set()
        for input_address in self.inputs:
            child = self.cellmap[input_address]
            if child == None:
                print("Not found ", input_address)
                continue
            g = make_subgraph(G, child, "descending")
            dependencies = dependencies.union(g.nodes())

        # print "%s cells depending on inputs" % str(len(dependencies))

        # prune the graph and set all cell independent of input to const
        subgraph = networkx.DiGraph()
        new_cellmap = {}
        for output_address in self.outputs:
            new_cellmap[output_address] = self.cellmap[output_address]
            seed = self.cellmap[output_address]
            todo = [(seed,n) for n in G.predecessors(seed)]
            done = set(todo)

            while len(todo) > 0:
                current, pred = todo.pop()
                # print "==========================="
                # print current.address(), pred.address()
                if current in dependencies:
                    if pred in dependencies or isinstance(pred.value, RangeCore) or pred.is_named_range:
                        subgraph.add_edge(pred, current)
                        new_cellmap[pred.address()] = pred
                        new_cellmap[current.address()] = current

                        nexts = G.predecessors(pred)
                        for n in nexts:
                            if (pred,n) not in done:
                                todo += [(pred,n)]
                                done.add((pred,n))
                    else:
                        if pred.address() not in new_cellmap:
                            const_node = Cell(pred.address(), pred.sheet, value = pred.range if pred.is_range else pred.value, formula=None, is_range = isinstance(pred.range, RangeCore), is_named_range=pred.is_named_range, should_eval=pred.should_eval)
                            # pystr,ast = cell2code(self.named_ranges, const_node, pred.sheet)
                            # const_node.python_expression = pystr
                            # const_node.compile()
                            new_cellmap[pred.address()] = const_node

                        const_node = new_cellmap[pred.address()]
                        subgraph.add_edge(const_node, current)

                else:
                    # case of range independant of input, we add all children as const
                    if pred.address() not in new_cellmap:
                        const_node = Cell(pred.address(), pred.sheet, value = pred.range if pred.is_range else pred.value, formula=None, is_range = pred.is_range, is_named_range=pred.is_named_range, should_eval=pred.should_eval)
                        # pystr,ast = cell2code(self.named_ranges, const_node, pred.sheet)
                        # const_node.python_expression = pystr
                        # const_node.compile()
                        new_cellmap[pred.address()] = const_node

                    const_node = new_cellmap[pred.address()]
                    subgraph.add_edge(const_node, current)


        print("Graph pruning done, %s nodes, %s edges, %s cellmap entries" % (len(subgraph.nodes()),len(subgraph.edges()),len(new_cellmap)))
        undirected = networkx.Graph(subgraph)
        # print "Number of connected components %s", str(number_connected_components(undirected))
        # print map(lambda x: x.address(), subgraph.nodes())

        # add back inputs that have been pruned because they are outside of calculation chain
        for i in self.inputs:
            if i not in new_cellmap:
                if i in self.named_ranges:
                    reference = self.named_ranges[i]
                    if is_range(reference):

                        rng = self.Range(reference)
                        virtual_cell = Cell(i, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                        new_cellmap[i] = virtual_cell
                        subgraph.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain

                    else:
                        # might need to be changed to actual self.cells Cell, not a copy
                        virtual_cell = Cell(i, None, value = self.cellmap[reference].value, formula = reference, is_range = False, is_named_range = True)
                        new_cellmap[i] = virtual_cell
                        subgraph.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                else:
                    if is_range(i):
                        rng = self.Range(i)
                        virtual_cell = Cell(i, None, value = rng, formula = o, is_range = True, is_named_range = True )
                        new_cellmap[i] = virtual_cell
                        subgraph.add_node(virtual_cell) # edges are not needed here since the input here is not in the calculation chain
                    else:
                        new_cellmap[i] = self.cellmap[i]
                        subgraph.add_node(self.cellmap[i]) # edges are not needed here since the input here is not in the calculation chain


        spreadsheet = Spreadsheet()
        return spreadsheet.build_spreadsheet(subgraph, new_cellmap, self.named_ranges, self.pointers, self.outputs, self.inputs, debug = self.debug)

    def clean_pointer(self):
        print('___### Cleaning Pointers ###___')

        new_named_ranges = self.named_ranges.copy()
        new_cells = self.cellmap.copy()

        ### 1) create ranges
        for n in self.named_ranges:
            reference = self.named_ranges[n]
            if is_range(reference):
                if 'OFFSET' not in reference:
                    my_range = self.Range(reference)
                    self.cellmap[n] = Cell(n, None, value = my_range, formula = reference, is_range = True, is_named_range = True )
                else:
                    self.cellmap[n] = Cell(n, None, value = None, formula = reference, is_range = False, is_named_range = True )
            else:
                if reference in self.cellmap:
                    self.cellmap[n] = Cell(n, None, value = self.cellmap[reference].value, formula = reference, is_range = False, is_named_range = True )
                else:
                    self.cellmap[n] = Cell(n, None, value = None, formula = reference, is_range = False, is_named_range = True )

        ### 2) gather all occurence of pointer functions in cells or named_range
        all_pointers = set()

        for pointer_name in self.pointer_to_remove:
            for k,v in list(self.named_ranges.items()):
                if pointer_name in v:
                    all_pointers.add((v, k, None))
            for k,cell in list(self.cellmap.items()):
                if cell.formula and pointer_name in cell.formula:
                    all_pointers.add((cell.formula, cell.address(), cell.sheet))

            # print "%s %s to parse" % (str(len(all_pointers)), pointer_name)

        ### 3) evaluate all pointers

        for formula, address, sheet in all_pointers:

            if sheet:
                parsed = parse_cell_address(address)
            else:
                parsed = ""
            e = shunting_yard(formula, self.named_ranges, ref=parsed, tokenize_range = True)
            ast,root = build_ast(e)
            code = root.emit(ast)

            cell = {"formula": formula, "address": address, "sheet": sheet}

            replacements = self.eval_pointers_from_ast(ast, root, cell)

            new_formula = formula
            if type(replacements) == list:
                for repl in replacements:
                    if type(repl["value"]) == ExcelError:
                        if self.debug:
                            print('WARNING: Excel error found => replacing with #N/A')
                        repl["value"] = "#N/A"

                    if repl["expression_type"] == "value":
                        new_formula = new_formula.replace(repl["formula"], str(repl["value"]))
                    else:
                        new_formula = new_formula.replace(repl["formula"], repl["value"])
            else:
                new_formula = None

            if address in new_named_ranges:
                new_named_ranges[address] = new_formula
            else:
                old_cell = self.cellmap[address]
                new_cells[address] = Cell(old_cell.address(), old_cell.sheet, value=old_cell.value, formula=new_formula, is_range = old_cell.is_range, is_named_range=old_cell.is_named_range, should_eval=old_cell.should_eval)

        return new_cells, new_named_ranges

    def print_value_ast(self, ast,node,indent):
        print("%s %s %s %s" % (" "*indent, str(node.token.tvalue), str(node.token.ttype), str(node.token.tsubtype)))
        for c in node.children(ast):
            self.print_value_ast(ast, c, indent+1)

    def eval_pointers_from_ast(self, ast, node, cell):
        results = []
        context = cell["sheet"]

        if (node.token.tvalue == "INDEX" or node.token.tvalue == "OFFSET"):
            pointer_string = reverse_rpn(node, ast)
            expression = node.emit(ast, context=context)

            if expression.startswith("self.eval_ref"):
                expression_type = "value"
            else:
                expression_type = "formula"

            try:
                pointer_value = eval(expression)

            except Exception as e:
                if self.debug:
                    print('EXCEPTION raised in eval_pointers: EXPR', expression, cell["address"])
                raise Exception("Problem evalling: %s for %s, %s" % (e, cell["address"], expression))

            return {"formula":pointer_string, "value": pointer_value, "expression_type": expression_type}
        else:
            for c in node.children(ast):
                results.append(self.eval_pointers_from_ast(ast, c, cell))
        return list(flatten(results, only_lists = True))


    def detect_alive(self, inputs = None, outputs = None):

        pointer_arguments = self.find_pointer_arguments(outputs)

        if inputs is None:
            inputs = self.inputs

        # go down the tree and list all cells that are pointer arguments
        todo = [self.cellmap[input] for input in inputs]
        done = set()
        alive = set()

        while len(todo) > 0:
            cell = todo.pop()

            if cell not in done:
                if cell.address() in pointer_arguments:
                    alive.add(cell.address())

                for child in self.G.successors(cell):
                    todo.append(child)

                done.add(cell)

        self.pointers_to_reset = alive
        return alive


    def find_pointer_arguments(self, outputs = None):

        # 1) gather all occurence of pointer
        all_pointers = set()

        if outputs is None:
            # 1.1) from all cells
            for pointer_name in self.pointer_to_remove:
                for k, cell in list(self.cellmap.items()):
                    if cell.formula and pointer_name in cell.formula:
                        all_pointers.add((cell.formula, cell.address(), cell.sheet))

        else:
            # 1.2) from the outputs while climbing up the tree
            todo = [self.cellmap[output] for output in outputs]
            done = set()
            while len(todo) > 0:
                cell = todo.pop()

                if cell not in done:
                    if cell.address() in self.pointers:
                        if cell.formula:
                            all_pointers.add((cell.formula, cell.address(), cell.sheet if cell.sheet is not None else None))
                        else:
                            raise Exception('Volatiles should always have a formula')

                    for parent in self.G.predecessors(cell): # climb up the tree
                        todo.append(parent)

                    done.add(cell)

        # 2) extract the arguments from these pointers
        done = set()
        pointer_arguments = set()

        #print 'All vol %i / %i' % (len(all_pointers), len(self.pointers))

        for formula, address, sheet in all_pointers:
            if formula not in done:
                if sheet:
                    parsed = parse_cell_address(address)
                else:
                    parsed = ""
                e = shunting_yard(formula, self.named_ranges, ref=parsed, tokenize_range = True)
                ast,root = build_ast(e)
                code = root.emit(ast)

                for a in list(flatten(self.get_pointer_arguments_from_ast(ast, root, sheet))):
                    pointer_arguments.add(a)

                done.add(formula)

        return pointer_arguments


    def get_arguments_from_ast(self, ast, node, sheet):
        arguments = []

        for c in node.children(ast):
            if c.tvalue == ":":
                arg_range =  reverse_rpn(c, ast)
                for elem in resolve_range(arg_range, False, sheet)[0]:
                    arguments += [elem]
            if c.ttype == "operand":
                if not is_number(c.tvalue):
                    if sheet is not None and "!" not in c.tvalue and c.tvalue not in self.named_ranges:
                        arguments += [sheet + "!" + c.tvalue]
                    else:
                        arguments += [c.tvalue]
            else:
                arguments += [self.get_arguments_from_ast(ast, c, sheet)]

        return arguments

    def get_pointer_arguments_from_ast(self, ast, node, sheet):
        arguments = []

        if node.token.tvalue in self.pointer_to_remove:
            for c in node.children(ast)[1:]:
                if c.ttype == "operand":
                    if not is_number(c.tvalue):
                        if sheet is not None and "!" not in c.tvalue and c.tvalue not in self.named_ranges:
                            arguments += [sheet + "!" + c.tvalue]
                        else:
                            arguments += [c.tvalue]
                else:
                        arguments += [self.get_arguments_from_ast(ast, c, sheet)]
        else:
            for c in node.children(ast):
                arguments += [self.get_pointer_arguments_from_ast(ast, c, sheet)]

        return arguments


    def dump_json(self, fname):
        dump_json(self, fname)

    def dump(self, fname):
        dump(self, fname)

    @staticmethod
    def load(fname):
        spreadsheet = Spreadsheet()
        spreadsheet.build_spreadsheet(*load(fname))
        return spreadsheet

    @staticmethod
    def load_json(fname):
        data = load_json(fname)
        return Spreadsheet.from_dict(data)

    def set_value(self, address, val):
        # previously set_value was used. Capture this behaviour.
        warnings.warn(
            "This function is depricated and will be replaced by cell_set_value. Please use this function instead. "
            "This behaviour will be removed in a future version.",
            PendingDeprecationWarning
        )
        return self.cell_set_value(address, val)

    def cell_set_value(self, address, value):
        """
        Set the value of a cell

        :param address: the address of a cell
        :param value: the new value
        """
        self.reset_buffer = set()

        try:
            address = address.replace('$', '')
            address = address.replace("'", '')
            cell = self.cellmap[address]

            # when you set a value on cell, its should_eval flag is set to 'never' so its formula is not used until set free again => sp.activate_formula()
            self.fix_cell(address)

            # case where the address refers to a range
            if cell.is_range:
                cells_to_set = []

                if not isinstance(value, list):
                    value = [value] * len(cells_to_set)

                self.cell_reset(cell.address())
                cell.range.values = value

            # case where the address refers to a single value
            else:
                if address in self.named_ranges:  # if the cell is a named range, we need to update and fix the reference cell
                    ref_address = self.named_ranges[address]

                    if ref_address in self.cellmap:
                        ref_cell = self.cellmap[ref_address]
                    else:
                        ref_cell = Cell(
                            ref_address, None, value=value,
                            formula=None, is_range=False, is_named_range=False)
                        self.cell_add(cell=ref_cell)

                    ref_cell.value = value

                if cell.value != value:
                    if cell.value is None:
                        cell.value = 'notNone'  # hack to avoid the direct return in reset() when value is None
                    # reset the node + its dependencies
                    self.cell_reset(cell.address())
                    # set the value
                    cell.value = value

            for vol in self.pointers_to_reset:  # reset all pointers
                self.cell_reset(self.cellmap[vol].address())
        except KeyError:
            raise Exception('Cell %s not in cellmap' % address)

    def reset(self, depricated=None):
        """
        Resets all the cells in a spreadsheet and indicates that an update is required.

        :return: nothing
        """

        # previously reset was used to only reset one cell. Capture this behaviour.
        if depricated is not None:
            warnings.warn(
                "reset() is used to reset the full spreadsheet, cell_reset() should be used to reset only one cell. "
                "This behaviour will be removed in a future version.",
                PendingDeprecationWarning
            )
            self.cell_reset(depricated.address())

        for cell in self.cellmap.values:
            self.cell_reset(cell.address())
        return

    def cell_reset(self, address):
        """
        Resets the value of the cell and indicates that an update is required. Also resets all of its dependents.

        :param address: the address of the cell to be reset.
        :return: nothing
        """

        if address in self.cellmap:
            cell = self.cellmap[address]
        else:
            return
        if cell.value is None and address not in self.named_ranges:
            return

        # check if cell has to be reset
        if cell.value is None and cell.need_update:
            return

        # update cells
        if cell.should_eval != 'never':
            if not cell.is_range:
                cell.value = None

            self.reset_buffer.add(cell)
            cell.need_update = True

        for child in self.G.successors(cell):
            if child not in self.reset_buffer:
                self.cell_reset(child.address())

    def fix_cell(self, address):
        warnings.warn(
            "xxx_cell functions are depricated and replaced by cell_xxx functions. Please use those functions instead. "
            "This behaviour will be removed in a future version.",
            PendingDeprecationWarning
        )
        return self.cell_fix(address)

    def cell_fix(self, address):
        """
        Fix the value of a cell

        :param address: the address of the cell
        """
        try:
            if address not in self.fixed_cells:
                cell = self.cellmap[address]
                self.fixed_cells[address] = cell.should_eval
                cell.should_eval = 'never'
        except KeyError:
            raise Exception('Cell %s not in cellmap' % address)

    def free_cell(self, address=None):
        warnings.warn(
            "xxx_cell functions are depricated and replaced by cell_xxx functions. Please use those functions instead. "
            "This behaviour will be removed in a future version.",
            PendingDeprecationWarning
        )
        return self.cell_free(address)

    def cell_free(self, address=None):
        """
        Free the cell (opposite of fix)

        :param address: the address of the cell
        """
        if address is None:
            for addr in self.fixed_cells:
                cell = self.cellmap[addr]

                cell.should_eval = 'always' # this is to be able to correctly reinitiliaze the value
                if cell.python_expression is not None:
                    self.eval_ref(addr)

                cell.should_eval = self.fixed_cells[addr]
            self.fixed_cells = {}

        else:
            try:
                cell = self.cellmap[address]

                cell.should_eval = 'always' # this is to be able to correctly reinitiliaze the value
                if cell.python_expression is not None:
                    self.eval_ref(address)

                cell.should_eval = self.fixed_cells[address]
                self.fixed_cells.pop(address, None)
            except KeyError:
                raise Exception('Cell %s not in cellmap' % address)

    def print_value_tree(self,addr,indent):
        cell = self.cellmap[addr]
        print("%s %s = %s" % (" "*indent,addr,cell.value))
        for c in self.G.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    def build_pointer(self, pointer):
        if not isinstance(pointer, RangeCore):
            vol_range = self.cellmap[pointer].range
        else:
            vol_range = pointer

        start = eval(vol_range.reference['start'])
        end = eval(vol_range.reference['end'])

        vol_range.build('%s:%s' % (start, end), debug = True)


    def build_pointers(self):

        for pointer in self.pointers:
            vol_range = self.cellmap[pointer].range

            start = eval(vol_range.reference['start'])
            end = eval(vol_range.reference['end'])

            vol_range.build('%s:%s' % (start, end), debug = True)

    def eval_ref(self, addr1, addr2 = None, ref = None):
        debug = False

        if isinstance(addr1, ExcelError):
            return addr1
        elif isinstance(addr2, ExcelError):
            return addr2
        else:
            if addr1 in self.cellmap:
                cell1 = self.cellmap[addr1]
            else:
                if self.debug:
                    print('WARNING in eval_ref: address %s not found in cellmap, returning #NULL' % addr1)
                return ExcelError('#NULL', 'Cell %s is empty' % addr1)
            if addr2 == None:
                if cell1.is_range:

                    if cell1.range.is_pointer:
                        self.build_pointer(cell1.range)
                        # print 'NEED UPDATE', cell1.need_update

                    associated_addr = RangeCore.find_associated_cell(ref, cell1.range)

                    if associated_addr: # if range is associated to ref, no need to return/update all range
                        return self.evaluate(associated_addr)
                    else:
                        range_name = cell1.address()
                        if cell1.need_update:
                            self.update_range(cell1.range)

                            range_need_update = True
                            for c in self.G.successors(cell1): # if a parent doesnt need update, then cell1 doesnt need update
                                if not c.need_update:
                                    range_need_update = False
                                    break

                            cell1.need_update = range_need_update
                            return cell1.range
                        else:
                            return cell1.range

                elif addr1 in self.named_ranges or not is_range(addr1):
                    value = self.evaluate(addr1)
                    return value
                else: # addr1 = Sheet1!A1:A2 or Sheet1!A1:Sheet1!A2
                    addr1, addr2 = addr1.split(':')
                    if '!' in addr1:
                        sheet = addr1.split('!')[0]
                    else:
                        sheet = None
                    if '!' in addr2:
                        addr2 = addr2.split('!')[1]

                    return self.Range('%s:%s' % (addr1, addr2))
            else:  # addr1 = Sheet1!A1, addr2 = Sheet1!A2
                if '!' in addr1:
                    sheet = addr1.split('!')[0]
                else:
                    sheet = None
                if '!' in addr2:
                    addr2 = addr2.split('!')[1]
                return self.Range('%s:%s' % (addr1, addr2))

    def update_range(self, range):
        # This function loops through its Cell references to evaluate the ones that need so
        # This uses Spreadsheet.pending dictionary, that holds the addresses of the Cells that are being calculated

        debug = False

        for index, key in enumerate(range.order):
            addr = get_cell_address(range.sheet, key)

            if self.cellmap[addr].need_update or self.cellmap[addr].value is None:
                self.evaluate(addr)

    def evaluate(self, cell, is_addr=True):
        if isinstance(cell, Cell):
            is_addr = False

        if is_addr:
            address = cell
        else:
            address = cell.address
        return self.cell_evaluate(address)

    def cell_evaluate(self, address):
        """
        Evaluate the cell.

        :param address: the address of the cell
        :return:
        """
        try:
            cell = self.cellmap[address]
        except:
            if self.debug:
                print('WARNING: Empty cell at ' + address)
            return ExcelError('#NULL', 'Cell %s is empty' % address)

        # no formula, fixed value
        if cell.should_eval == 'normal' and not cell.need_update and cell.value is not None or not cell.formula or cell.should_eval == 'never':
            return cell.value if cell.value != '' else None
        try:
            if cell.is_range:
                for child in cell.range.cells:
                    self.evaluate(child.address())
            elif cell.compiled_expression != None:
                vv = eval(cell.compiled_expression)
                if isinstance(vv, RangeCore): # this should mean that vv is the result of RangeCore.apply_all, but with only one value inside
                    cell.value = vv.values[0]
                else:
                    cell.value = vv if vv != '' else None
            else:
                cell.value = 0

            cell.need_update = False

            # DEBUG: saving differences
            if self.save_history:
                if cell.address() in self.history:
                    ori_value = self.history[cell.address()]['original']

                    if 'new' not in list(self.history[cell.address()].keys()):
                        if type(ori_value) == list and type(cell.value) == list \
                                and all([not is_almost_equal(x_y[0], x_y[1]) for x_y in zip(ori_value, cell.value)]) \
                                or not is_almost_equal(ori_value, cell.value):

                            self.count += 1
                            self.history[cell.address()]['formula'] = str(cell.formula)
                            self.history[cell.address()]['priority'] = self.count
                            self.history[cell.address()]['python'] = str(cell.python_expression)

                            if self.count == 1:
                                self.history['ROOT_DIFF'] = self.history[cell.address()]
                                self.history['ROOT_DIFF']['cell'] = cell.address()

                    self.history[cell.address()]['new'] = str(cell.value)
                else:
                    if isinstance(cell.value, ExcelError):
                        self.history[cell.address()] = {'new': str(cell.value), 'error': str(cell.value.info)}
                    else:
                        self.history[cell.address()] = {'new': str(cell.value)}

        except Exception as e:
            if str(e).startswith("Problem evalling"):
                raise e
            else:
                raise Exception("Problem evalling: %s for %s, %s" % (e,cell.address(),cell.python_expression))

        return cell.value

    def asdict(self):
        data = json_graph.node_link_data(self.G)

        def cell_to_dict(cell):
            if isinstance(cell.range, RangeCore):
                range = cell.range
                value = {
                    "cells": range.addresses,
                    "values": range.values,
                    "nrows": range.nrows,
                    "ncols": range.ncols
                }
            else:
                value = cell.value

            node = {
                "address": cell.address(),
                "formula": cell.formula,
                "value": value,
                "python_expression": cell.python_expression,
                "is_named_range": cell.is_named_range,
                "should_eval": cell.should_eval
            }
            return node

        # save nodes as simple objects
        nodes = []
        for node in data["nodes"]:
            cell = node["id"]
            nodes.append(cell.asdict())

        links = []
        for el in data['links']:
            link = {key: cell.address() for key, cell in el.items()}
            links.append(link)

        data["nodes"] = nodes
        data["links"] = links
        data["outputs"] = self.outputs
        data["inputs"] = self.inputs
        data["named_ranges"] = self.named_ranges

        return data

    @staticmethod
    def from_dict(input_data):

        data = dict(input_data)

        nodes = list(
            map(Cell.from_dict,
                filter(
                    lambda item: not isinstance(item['value'], dict),
                    data['nodes'])))
        cellmap = {n.address(): n for n in nodes}

        def cell_from_dict(d):
            return Cell.from_dict(d, cellmap=cellmap)

        nodes.extend(
            list(
                map(cell_from_dict,
                    filter(
                        lambda item: isinstance(item['value'], dict),
                        data['nodes']))))

        data["nodes"] = [{'id': node} for node in nodes]

        links = []
        idmap = { node.address(): node for node in nodes }
        for el in data['links']:
            source_address = el['source']
            target_address = el['target']
            link = {
                'source': idmap[source_address],
                'target': idmap[target_address],
            }
            links.append(link)

        data['links'] = links

        G = json_graph.node_link_graph(data)
        cellmap = {n.address(): n for n in G.nodes()}

        named_ranges = data["named_ranges"]
        inputs = data["inputs"]
        outputs = data["outputs"]

        spreadsheet = Spreadsheet()
        spreadsheet.build_spreadsheet(
            G, cellmap, named_ranges,
            inputs=inputs, outputs=outputs)
        return spreadsheet
