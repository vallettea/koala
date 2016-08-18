# cython: profile=True

import os.path
import textwrap
from math import *

import networkx
from networkx.algorithms import number_connected_components

import json
import re

from koala.openpyxl.formula.translate import Translator

from koala.excellib import *
from koala.utils import *
from koala.ast import *
from koala.Cell import Cell
from koala.Range import RangeCore, RangeFactory, parse_cell_address, get_cell_address
from koala.tokenizer import reverse_rpn
from koala.serializer import *

class Spreadsheet(object):
    def __init__(self, G, cellmap, named_ranges, volatiles = set(), outputs = set(), inputs = set(), debug = False):
        super(Spreadsheet,self).__init__()
        self.G = G
        self.cellmap = cellmap
        self.named_ranges = named_ranges

        addr_to_name = {}
        for name in named_ranges:
            addr_to_name[named_ranges[name]] = name
        self.addr_to_name = addr_to_name

        addr_to_range = {}        
        for c in self.cellmap.values():
            if c.is_range and len(c.range.keys()) != 0: # could be better, but can't check on Exception types here...
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
        self.volatile_to_remove = ["INDEX", "OFFSET"]
        self.volatiles = volatiles
        self.Range = RangeFactory(cellmap)
        self.reset_buffer = set()
        self.debug = debug
        self.pending = {}
        self.fixed_cells = {}

    def activate_history(self):
        self.save_history = True

    def add_cell(self, cell, value = None):
        
        if type(cell) != Cell:
            cell = Cell(cell, None, value = value, formula = None, is_range = False, is_named_range = False)
        
        addr = cell.address()
        if addr in self.cellmap:
            raise Exception('Cell %s already in cellmap' % addr)

        cellmap, G = graph_from_seeds([cell], self)

        self.cellmap = cellmap
        self.G = G

        print "Graph construction updated, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))

    def set_formula(self, addr, formula):
        if addr in self.cellmap:
            cell = self.cellmap[addr]
        else:
            raise Exception('Cell %s not in cellmap' % addr)

        seeds = [cell]

        if cell.is_range:
            for index, c in enumerate(cell.range.cells): # for each cell of the range, translate the formula
                if index == 0:
                    c.formula = formula
                    translator = Translator(unicode('=' + formula), c.address().split('!')[1]) # the Translator needs a reference without sheet
                else:
                    translated = translator.translate_formula(c.address().split('!')[1]) # the Translator needs a reference without sheet
                    c.formula = translated[1:] # to get rid of the '='

                seeds.append(c)
        else:
            cell.formula = formula

        cellmap, G = graph_from_seeds(seeds, self)

        self.cellmap = cellmap
        self.G = G

        should_eval = self.cellmap[addr].should_eval
        self.cellmap[addr].should_eval = 'always'
        self.evaluate(addr)
        self.cellmap[addr].should_eval = should_eval

        print "Graph construction updated, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))


    def prune_graph(self):
        print '___### Pruning Graph ###___'

        G = self.G

        # get all the cells impacted by inputs
        dependencies = set()
        for input_address in self.inputs:
            child = self.cellmap[input_address]
            if child == None:
                print "Not found ", input_address
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
            todo = map(lambda n: (seed,n), G.predecessors(seed))
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


        print "Graph pruning done, %s nodes, %s edges, %s cellmap entries" % (len(subgraph.nodes()),len(subgraph.edges()),len(new_cellmap))
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


        return Spreadsheet(subgraph, new_cellmap, self.named_ranges, self.volatiles, self.outputs, self.inputs, debug = self.debug)

    def clean_volatile(self, with_cache = True):
        print '___### Cleaning Volatiles ###___'
        # with_cache = False

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
        
        ### 2) gather all occurence of volatile functions in cells or named_range
        all_volatiles = set()

        for volatile_name in self.volatile_to_remove:
            for k,v in self.named_ranges.items():
                if volatile_name in v:
                    all_volatiles.add((v, k, None))
            for k,cell in self.cellmap.items():
                if cell.formula and volatile_name in cell.formula:
                    all_volatiles.add((cell.formula, cell.address(), cell.sheet))

            # print "%s %s to parse" % (str(len(all_volatiles)), volatile_name)

        ### 3) evaluate all volatiles
        if with_cache:
            cache = {} # formula => new_formula
        
        for formula, address, sheet in all_volatiles:

            if with_cache and formula in cache:
                # print 'Retrieving', cell["address"], cell["formula"], cache[cell["formula"]]
                new_formula = cache[formula]
            else:
                if sheet:
                    parsed = parse_cell_address(address)
                else:
                    parsed = ""
                e = shunting_yard(formula, self.named_ranges, ref=parsed, tokenize_range = True)
                ast,root = build_ast(e)
                code = root.emit(ast)
                
                cell = {"formula": formula, "address": address, "sheet": sheet}
                replacements = self.eval_volatiles_from_ast(ast, root, cell)

                new_formula = formula
                if type(replacements) == list:
                    for repl in replacements:
                        if type(repl["value"]) == ExcelError:
                            if self.debug:
                                print 'WARNING: Excel error found => replacing with #N/A'
                            repl["value"] = "#N/A"

                        if repl["expression_type"] == "value":
                            new_formula = new_formula.replace(repl["formula"], str(repl["value"]))
                        else:
                            new_formula = new_formula.replace(repl["formula"], repl["value"])
                else:
                    new_formula = None

                if with_cache:
                    # print 'Caching', cell["address"], cell["formula"], new_formula
                    cache[formula] = new_formula

            if address in new_named_ranges:
                new_named_ranges[address] = new_formula
            else: 
                old_cell = self.cellmap[address]
                new_cells[address] = Cell(old_cell.address(), old_cell.sheet, value=old_cell.value, formula=new_formula, is_range = old_cell.is_range, is_named_range=old_cell.is_named_range, should_eval=old_cell.should_eval)

        return new_cells, new_named_ranges

    def print_value_ast(self, ast,node,indent):
        print "%s %s %s %s" % (" "*indent, str(node.token.tvalue), str(node.token.ttype), str(node.token.tsubtype))
        for c in node.children(ast):
            self.print_value_ast(ast, c, indent+1)

    def eval_volatiles_from_ast(self, ast, node, cell):
        results = []
        context = cell["sheet"]

        if (node.token.tvalue == "INDEX" or node.token.tvalue == "OFFSET"):
            volatile_string = reverse_rpn(node, ast)
            expression = node.emit(ast, context=context)

            if expression.startswith("self.eval_ref"):
                expression_type = "value"
            else:
                expression_type = "formula"
            
            try:
                volatile_value = eval(expression)
            except Exception as e:
                if self.debug:
                    print 'EXCEPTION raised in eval_volatiles: EXPR', expression, cell["address"]
                raise Exception("Problem evalling: %s for %s, %s" % (e, cell["address"], expression))

            return {"formula":volatile_string, "value": volatile_value, "expression_type": expression_type}      
        else:
            for c in node.children(ast):
                results.append(self.eval_volatiles_from_ast(ast, c, cell))
        return list(flatten(results, only_lists = True))


    def detect_alive(self, inputs = None, outputs = None):

        volatile_arguments = self.find_volatile_arguments(outputs)
        
        if inputs is None:
            inputs = self.inputs

        # go down the tree and list all cell that are volatile arguments
        todo = [self.cellmap[input] for input in inputs]
        done = set()
        alive = set()

        while len(todo) > 0:
            cell = todo.pop()

            if cell not in done:
                if cell.address() in volatile_arguments:
                    alive.add(cell.address())

                # for child in self.G.predecessors_iter(cell):
                for child in self.G.successors_iter(cell):
                    todo.append(child)
  
                done.add(cell)
        return alive


    def find_volatile_arguments(self, outputs = None):
        
        # 1) gather all occurence of volatile 
        all_volatiles = set()

        if outputs is None:
            # 1.1) from all cells
            for volatile_name in self.volatile_to_remove:
                for k, cell in self.cellmap.items():
                    if cell.formula and volatile_name in cell.formula:
                        all_volatiles.add((cell.formula, cell.address(), cell.sheet))

        else:
            # 1.2) from the outputs while climbing up the tree
            todo = [self.cellmap[output] for output in outputs]
            done = set()

            while len(todo) > 0:
                cell = todo.pop()

                if cell not in done:
                    if cell.address() in self.volatiles:
                        if cell.formula:
                            all_volatiles.add((cell.formula, cell.address(), cell.sheet if cell.sheet is not None else None))
                        else:
                            raise Exception('Volatiles should always have a formula')

                    done.add(cell)

        # 2) extract the arguments from these volatiles
        done = set()
        volatile_arguments = set()

        for formula, address, sheet in all_volatiles:
            if formula not in done:
                if sheet:
                    parsed = parse_cell_address(address)
                else:
                    parsed = ""
                e = shunting_yard(formula, self.named_ranges, ref=parsed, tokenize_range = True)
                ast,root = build_ast(e)
                code = root.emit(ast)
                
                for a in list(flatten(self.get_volatile_arguments_from_ast(ast, root, sheet))):
                    volatile_arguments.add(a)

                done.add(formula)
            
        return volatile_arguments


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

    def get_volatile_arguments_from_ast(self, ast, node, sheet):
        arguments = []

        if node.token.tvalue in self.volatile_to_remove:
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
                arguments += [self.get_volatile_arguments_from_ast(ast, c, sheet)]

        return arguments


    def dump_json(self, fname):
        dump_json(self, fname)

    def dump(self, fname):
        dump(self, fname)

    @staticmethod
    def load(fname):
        return Spreadsheet(*load(fname))

    @staticmethod
    def load_json(fname):
        return Spreadsheet(*load_json(fname))
    
    def set_value(self, address, val):

        self.reset_buffer = set()

        if address not in self.cellmap:
            raise Exception("Address not present in graph.")

        address = address.replace('$','')
        cell = self.cellmap[address]

        # when you set a value on cell, its should_eval flag is set to 'never' so its formula is not used until set free again => sp.activate_formula()
        self.fix_cell(address)

        # case where the address refers to a range
        if self.cellmap[address].is_range: 
            cells_to_set = []
            # for a in self.cellmap[address].range.addresses:
                # if a in self.cellmap:
                #     cells_to_set.append(self.cellmap[a])
                #     self.fix_cell(a)

            if type(val) != list:
                val = [val]*len(cells_to_set)

            self.reset(cell)
            cell.range.values = val

        # case where the address refers to a single value
        else:
            if address in self.named_ranges: # if the cell is a named range, we need to update and fix the reference cell
                ref_address = self.named_ranges[address]
                
                if ref_address in self.cellmap:
                    ref_cell = self.cellmap[ref_address]
                else:
                    ref_cell = Cell(ref_address, None, value = val, formula = None, is_range = False, is_named_range = False )
                    self.add_cell(ref_cell)

                # self.fix_cell(ref_address)
                ref_cell.value = val

            if cell.value != val:
                if cell.value is None:
                    cell.value = 'notNone' # hack to avoid the direct return in reset() when value is None
                # reset the node + its dependencies
                self.reset(cell)
                # set the value
                cell.value = val

        for vol_range in self.volatiles: # reset all volatiles
            self.reset(self.cellmap[vol_range])

    def reset(self, cell):
        addr = cell.address()
        if cell.value is None and addr not in self.named_ranges: return

        # update cells
        if cell.should_eval != 'never':
            if not cell.is_range:
                cell.value = None

            self.reset_buffer.add(cell)
            cell.need_update = True

        for child in self.G.successors_iter(cell):
            if child not in self.reset_buffer:
                self.reset(child)

    def fix_cell(self, address):
        if address in self.cellmap:
            if address not in self.fixed_cells:
                cell = self.cellmap[address]
                self.fixed_cells[address] = cell.should_eval
                cell.should_eval = 'never'
        else:
            raise Exception('Cell %s not in cellmap' % address)

    def free_cell(self, address = None):
        if address is None:
            for addr in self.fixed_cells:
                cell = self.cellmap[addr]

                cell.should_eval = 'always' # this is to be able to correctly reinitiliaze the value
                if cell.python_expression is not None:
                    self.eval_ref(addr)
                
                cell.should_eval = self.fixed_cells[addr]
            self.fixed_cells = {}

        elif address in self.cellmap:
            cell = self.cellmap[address]

            cell.should_eval = 'always' # this is to be able to correctly reinitiliaze the value
            if cell.python_expression is not None:
                self.eval_ref(address)
            
            cell.should_eval = self.fixed_cells[address]
            self.fixed_cells.pop(address, None)
        else:
            raise Exception('Cell %s not in cellmap' % address)

    def print_value_tree(self,addr,indent):
        cell = self.cellmap[addr]
        print "%s %s = %s" % (" "*indent,addr,cell.value)
        for c in self.G.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    def build_volatile(self, volatile):
        if not isinstance(volatile, RangeCore):
            vol_range = self.cellmap[volatile].range
        else:
            vol_range = volatile

        start = eval(vol_range.reference['start'])
        end = eval(vol_range.reference['end'])

        vol_range.build('%s:%s' % (start, end), debug = True)

    def build_volatiles(self):

        for volatile in self.volatiles:
            vol_range = self.cellmap[volatile].range

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
                    print 'WARNING in eval_ref: address %s not found in cellmap, returning #NULL' % addr1
                return ExcelError('#NULL', 'Cell %s is empty' % addr1)
            if addr2 == None:
                if cell1.is_range:

                    if cell1.range.is_volatile:
                        self.build_volatile(cell1.range)

                    associated_addr = RangeCore.find_associated_cell(ref, cell1.range)

                    if associated_addr: # if range is associated to ref, no need to return/update all range
                        return self.evaluate(associated_addr)
                    else:
                        range_name = cell1.address()

                        if cell1.need_update:
                            self.update_range(cell1.range)
                            range_need_update = True
                            
                            for c in self.G.successors_iter(cell1): # if a parent doesnt need update, then cell1 doesnt need update
                                if not c.need_update:
                                    range_need_update = False
                                    break

                            cell1.need_update = range_need_update
                            return cell1.range
                        else:
                            return cell1.range

                elif addr1 in self.named_ranges or not is_range(addr1):
                    val = self.evaluate(addr1)
                    return val
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

        if range.name not in self.pending.keys():
            self.pending[range.name] = []

        for index, key in enumerate(range.order):
            addr = get_cell_address(range.sheet, key)
            if addr not in self.pending[range.name]:
                self.pending[range.name].append(addr)

                if self.cellmap[addr].need_update:
                    new_value = self.evaluate(addr)

        self.pending[range.name] = []
            

    def evaluate(self,cell,is_addr=True):
        if is_addr:
            try:
                cell = self.cellmap[cell]
            except:
                if self.debug:
                    print 'WARNING: Empty cell at ' + cell
                return ExcelError('#NULL', 'Cell %s is empty' % cell)    

        # no formula, fixed value
        if cell.should_eval == 'normal' and not cell.need_update and cell.value is not None or not cell.formula or cell.should_eval == 'never':
            return cell.value
        try:
            if cell.compiled_expression != None:
                vv = eval(cell.compiled_expression)
                if isinstance(vv, RangeCore): # this should mean that vv is the result of RangeCore.apply_all, but with only one value inside
                    cell.value = vv.values[0]
                else:
                    cell.value = vv
            elif cell.is_range:
                for child in cell.range.cells:
                    self.evaluate(child.address())
            else:
                cell.value = 0
            
            
            cell.need_update = False
            
            # DEBUG: saving differences
            if self.save_history:

                def is_almost_equal(a, b, precision = 0.001):
                    if is_number(a) and is_number(b):
                        return abs(float(a) - float(b)) <= precision
                    elif (a is None or a == 'None') and (b is None or b == 'None'):
                        return True
                    else:
                        return a == b

                if cell.address() in self.history:
                    ori_value = self.history[cell.address()]['original']
                    
                    if 'new' not in self.history[cell.address()].keys():
                        if type(ori_value) == list and type(cell.value) == list \
                                and all(map(lambda (x, y): not is_almost_equal(x, y), zip(ori_value, cell.value))) \
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
                    self.history[cell.address()] = {'new': str(cell.value)}

        except Exception as e:
            if e.message is not None and e.message.startswith("Problem evalling"):
                raise e
            else:
                raise Exception("Problem evalling: %s for %s, %s" % (e,cell.address(),cell.python_expression)) 

        return cell.value
