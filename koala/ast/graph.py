import os.path
import textwrap
import koala.ast.excellib as excelfun
from koala.ast.excellib import *
from koala.ast.excelutils import *
from math import *

import networkx
from networkx.classes.digraph import DiGraph
from networkx.drawing.nx_pydot import write_dot
from networkx.drawing.nx_pylab import draw, draw_circular
from networkx.readwrite.gexf import write_gexf
from networkx.readwrite import json_graph
from networkx.algorithms import number_connected_components

import networkx as nx

from astutils import find_node, subgraph

from tokenizer import ExcelParser, f_token, shunting_yard, reverse_rpn
import cPickle
import logging
from itertools import chain

from Range import RangeCore, RangeFactory, parse_cell_address, get_cell_address

import json
import gzip

from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from ..excel.utils import rows_from_range
from ExcelError import ExcelError, EmptyCellError, ErrorCodes



class Spreadsheet(object):
    def __init__(self, G, cellmap, named_ranges, outputs = [],  inputs = []):
        super(Spreadsheet,self).__init__()
        self.G = G
        self.cellmap = cellmap
        self.named_ranges = named_ranges
        self.outputs = outputs
        self.inputs = inputs
        self.history = dict()
        self.count = 0
        self.volatile_to_remove = ["INDEX", "OFFSET"]
        self.Range = RangeFactory(cellmap)
        self.reset_buffer = set()

    def prune_graph(self, inputs):
        print '___### Pruning Graph ###___'

        G = self.G

        # get all the cells impacted by inputs
        dependencies = set()
        for input_address in inputs:
                child = find_node(G, input_address)
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
            seed = find_node(G, output_address)
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
                            const_node = Cell(pred.address(), pred.sheet, value=pred.value, formula=None, is_range = isinstance(pred.value, RangeCore), is_named_range=False, always_eval=False)
                            # pystr,ast = cell2code(self.named_ranges, const_node, pred.sheet)
                            # const_node.python_expression = pystr
                            # const_node.compile()
                            new_cellmap[pred.address()] = const_node

                        const_node = new_cellmap[pred.address()]
                        subgraph.add_edge(const_node, current)
                        
                else:
                    # case of range independant of input, we add all children as const
                    if pred.address() not in new_cellmap:
                        const_node = Cell(pred.address(), pred.sheet, value=pred.value, formula=None, is_range = isinstance(pred.value, RangeCore), is_named_range=False, always_eval=False)
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

        return Spreadsheet(subgraph, new_cellmap, self.named_ranges, self.outputs, inputs)

    def clean_volatile(self):

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
        all_volatiles = []

        for volatile_name in self.volatile_to_remove:
            for k,v in self.named_ranges.items():
                if volatile_name in v:
                    all_volatiles.append({"formula":v, "address": k, "sheet": None})
            for k,cell in self.cellmap.items():
                if cell.formula and volatile_name in cell.formula:
                    all_volatiles.append({"formula":cell.formula, "address": cell.address(), "sheet": cell.sheet})

            # print "%s %s to parse" % (str(len(all_volatiles)), volatile_name)

        ### 3) evaluate all volatiles
        cache = {} # formula => new_formula

        for cell in all_volatiles:

            if cell["formula"] in cache:
                new_formula = cache[cell["formula"]]
            else:
                if cell["sheet"]:
                    parsed = parse_cell_address(cell["address"])
                else:
                    parsed = ""
                e = shunting_yard(cell["formula"], self.named_ranges, ref=parsed, tokenize_range = True)
                ast,root = build_ast(e)
                code = root.emit(ast)
                
                replacements = self.eval_volatiles_from_ast(ast, root, cell)

                new_formula = cell["formula"]
                if type(replacements) == list:
                    for repl in replacements:
                        if type(repl["value"]) == ExcelError:
                            print 'EXCEL ERROR ENCOUNTERED', repl["value"].value, repl["value"].info
                            repl["value"] = "#N/A"

                        if repl["expression_type"] == "value":
                            new_formula = new_formula.replace(repl["formula"], str(repl["value"]))
                        else:
                            new_formula = new_formula.replace(repl["formula"], repl["value"])
                else:
                    new_formula = None
                cache[cell["formula"]] = new_formula

            if cell["address"] in new_named_ranges:
                new_named_ranges[cell["address"]] = new_formula
            else: 
                old_cell = self.cellmap[cell["address"]]
                new_cells[cell["address"]] = Cell(old_cell.address(), old_cell.sheet, value=old_cell.value, formula=new_formula, is_range = old_cell.is_range, is_named_range=old_cell.is_named_range, always_eval=old_cell.always_eval)
            
        return new_cells, new_named_ranges

    def print_value_ast(self, ast,node,indent):
        print "%s %s %s %s" % (" "*indent, str(node.token.tvalue), str(node.token.ttype), str(node.token.tsubtype))
        for c in node.children(ast):
            self.print_value_ast(ast, c, indent+1)

    def eval_volatiles_from_ast(self, ast, node, cell):
        results = []
        context = cell["sheet"]

        if (node.token.tvalue == "INDEX" and node.parent(ast) is not None and node.parent(ast).tvalue == ':') or \
            (node.token.tvalue == "OFFSET"):
            #print self.print_value_ast(ast, node, 1)
            volatile_string = reverse_rpn(node, ast)
            expression = node.emit(ast, context=context)
            #print expression
            #print '->', volatile_string
            if expression.startswith("self.eval_ref"):
                expression_type = "value"
            else:
                expression_type = "formula"
            
            try:
                volatile_value = eval(expression)
            except Exception as e:
                raise Exception("Problem evalling: %s for %s, %s" % (e, cell["address"], expression)) 

            return {"formula":volatile_string, "value": volatile_value, "expression_type": expression_type}      
        else:
            for c in node.children(ast):
                results.append(self.eval_volatiles_from_ast(ast, c, cell))
        return list(flatten(results, only_lists = True))


    def dump(self, fname):
        data = json_graph.node_link_data(self.G)
        # save nodes as simple objects
        nodes = []
        for node in data["nodes"]:
            cell = node["id"]

            if isinstance(cell.range, RangeCore):
                range = cell.range
                value = {
                    "cells": range.cells,
                    "values": range.value,
                    "nrows": range.nrows,
                    "ncols": range.ncols,
                    "name": range.name
                }
            else:
                value = cell.value

            nodes += [{
                "address": cell.address(),
                "formula": cell.formula,
                "value": value,
                "python_expression": cell.python_expression,
                "is_named_range": cell.is_named_range,
                "always_eval": cell.always_eval
            }]
        data["nodes"] = nodes
        data["outputs"] = self.outputs
        data["inputs"] = self.inputs
        data["named_ranges"] = self.named_ranges
        with gzip.GzipFile(fname, 'w') as outfile:
            outfile.write(json.dumps(data))

    @staticmethod
    def load(fname):

        def _decode_list(data):
            rv = []
            for item in data:
                if isinstance(item, unicode):
                    item = item.encode('utf-8')
                elif isinstance(item, list):
                    item = _decode_list(item)
                elif isinstance(item, dict):
                    item = _decode_dict(item)
                rv.append(item)
            return rv

        def _decode_dict(data):
            rv = {}
            for key, value in data.iteritems():
                if isinstance(key, unicode):
                    key = key.encode('utf-8')
                if isinstance(value, unicode):
                    value = value.encode('utf-8')
                elif isinstance(value, list):
                    value = _decode_list(value)
                elif isinstance(value, dict):
                    value = _decode_dict(value)
                rv[key] = value
            return rv
        with gzip.GzipFile(fname, 'r') as infile:
            data = json.loads(infile.read(), object_hook=_decode_dict)
        def cell_from_dict(d):
            cell_is_range = type(d["value"]) == dict
            if cell_is_range:
                range = d["value"]
                if len(range["values"]) == 0:
                    range["values"] = [None] * len(range["cells"])
                value = RangeCore(range["cells"], range["values"], nrows = range["nrows"], ncols = range["ncols"], name = range["name"])
            else:
                value = d["value"]
            new_cell = Cell(d["address"], None, value=value, formula=d["formula"], is_range = cell_is_range, is_named_range=d["is_named_range"], always_eval=d["always_eval"])
            new_cell.python_expression = d["python_expression"]
            new_cell.compile()
            return {"id": new_cell}

        nodes = map(cell_from_dict, data["nodes"])
        data["nodes"] = nodes

        G = json_graph.node_link_graph(data)
        cellmap = {n.address():n for n in G.nodes()}

        return Spreadsheet(G, cellmap, data["named_ranges"], data["outputs"], data["inputs"])

    def export_to_dot(self,fname):
        write_dot(self.G,fname)
                    
    def export_to_gexf(self,fname):
        write_gexf(self.G,fname)
    
    def plot_graph(self):
        import matplotlib.pyplot as plt

        pos=nx.spring_layout(self.G,iterations=2000)
        #pos=nx.spectral_layout(G)
        #pos = nx.random_layout(G)
        nx.draw_networkx_nodes(self.G, pos)
        nx.draw_networkx_edges(self.G, pos, arrows=True)
        nx.draw_networkx_labels(self.G, pos)
        plt.show()
    
    def set_value(self, address, val):

        self.reset_buffer = set()

        if address not in self.cellmap:
            raise Exception("Address not present in graph.")

        # case where the address refers to a range

        if self.cellmap[address].range:
            cell_to_set = [self.cellmap[a] for a in self.cellmap[address].range.cells if a in self.cellmap]
            if type(val) != list:
                val = [val]*len(cell_to_set)
            for cell, value in zip(cell_to_set, val):
                if cell.value != value:
                    # reset the node + its dependencies
                    self.reset(cell)
                    # set the value
                    cell.value = value

        # case where the address refers to a single value
        else:
            address = address.replace('$','')
            cell = self.cellmap[address]
            if cell.value != val:
                # reset the node + its dependencies
                self.reset(cell)
                # set the value
                cell.value = val

    def reset(self, cell):
        addr = cell.address()
        if cell.value is None and addr not in self.named_ranges: return

        self.reset_buffer.add(cell)
        # update depending ranges
        if cell.range:
            cell.range.reset()
        else:
            cell.value = None
            cell.need_update = True
        for child in self.G.successors_iter(cell):
            if child not in self.reset_buffer:
                self.reset(child)

    def print_value_tree(self,addr,indent):
        cell = self.cellmap[addr]
        print "%s %s = %s" % (" "*indent,addr,cell.value)
        for c in self.G.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    ### might not be needed anymore
    
    # def recalculate(self):
    #     for c in self.cellmap.values():
    #         if isinstance(c,CellRange):
    #             self.evaluate_range(c,is_addr=False)
    #         else:
    #             self.evaluate(c,is_addr=False)
                
    # def evaluate_range(self,rng,is_addr=True):

    ### might not be needed anymore

    #     address = None
    #     if is_addr:
    #         address = rng
    #         if address in cellmap:
    #             return self.cellmap[address]
    #         else:
    #             return self.Range(address)
    #     else:
    #         address = rng.celladdrs

    #     # its important that [] gets treated ad false here
    #     if rng.value:
    #         return rng.value

    #     # cells,nrows,ncols = rng.celladdrs,rng.nrows,rng.ncols

    #     # cells = list(flatten(cells))

    #     # values = [ self.evaluate(c) for c in cells if c in self.cellmap]
    #     # cells = [c for c in cells if c in self.cellmap]

    #     data = self.Range()
    #     rng.value = data
        
    #     return data

    def eval_ref(self, addr1, addr2 = None):
        if isinstance(addr1, ExcelError):
            return addr1
        elif isinstance(addr2, ExcelError):
            return addr2
        else:
            try:
                cell1 = self.cellmap[addr1]
            except:
                print 'Eval_ref Warning: address %s not found in cellmap, returning #NULL' % addr1
                return ExcelError('#NULL', 'Cell %s is empty' % addr1)
            if addr2 == None:
                if cell1.is_range:
                    if cell1.range.need_update:
                        cell1.range.need_update = False
                        return self.update_range(cell1.range)
                    else:
                        return cell1.range
                elif addr1 in self.named_ranges or not is_range(addr1):
                    return self.evaluate(addr1)
                else: # addr1 = Sheet1!A1:A2 or Sheet1!A1:Sheet1!A2
                    # if addr1 == "Cashflow!L39:L50":

                    addr1, addr2 = addr1.split(':')
                    if '!' in addr1:
                        sheet = addr1.split('!')[0]
                    else:
                        sheet = None
                    if '!' in addr2:
                        addr2 = addr2.split('!')[1]

                    return self.Range('%s:%s' % (addr1, addr2))
                    # return self.evaluate_range(CellRange('%s:%s' % (addr1, addr2),sheet), False)
            else:  # addr1 = Sheet1!A1, addr2 = Sheet1!A2
                if '!' in addr1:
                    sheet = addr1.split('!')[0]
                else:
                    sheet = None
                if '!' in addr2:
                    addr2 = addr2.split('!')[1]
                return self.Range('%s:%s' % (addr1, addr2))
                # return self.evaluate_range(CellRange('%s:%s' % (addr1, addr2),sheet), False)

    def update_range(self, range):
        debug = False

        for key in range: # only ranges with need_update to True are updated, so all values are None and need evaluation
            addr = get_cell_address(range.sheet, key)
            range[key] = self.evaluate(addr)
            
        return range

    def evaluate(self,cell,is_addr=True):

        if is_addr:
            try:
                cell = self.cellmap[cell]
            except:
                # print 'Empty cell at '+ cell
                return ExcelError('#NULL', 'Cell %s is empty' % cell)

        # no formula, fixed value
        if not cell.formula or not cell.always_eval and not cell.need_update:
            # print "returning constant or cached value for ", cell.address()
            
            return cell.value
        
        try:
            # print "Evalling: %s, %s" % (cell.address(),cell.python_expression)
            if cell.compiled_expression != None:
                vv = eval(cell.compiled_expression)
            else:
                vv = 0
            cell.value = vv
            cell.need_update = False

            # # DEBUG: saving differences
            # if cell.address() in self.history:
            #     ori_value = self.history[cell.address()]['original']
            #     if is_number(ori_value) and is_number(cell.value) and abs(float(ori_value) - float(cell.value)) > 0.001:
            #         self.count += 1
            #         self.history[cell.address()]['formula'] = str(cell.formula)
            #         self.history[cell.address()]['priority'] = self.count
            #         self.history[cell.address()]['python'] = str(cell.python_expression)

            #     self.history[cell.address()]['new'] = str(cell.value)
            # else:
            #     self.history[cell.address()] = {'new': str(cell.value)}

        except Exception as e:
            if e.message.startswith("Problem evalling"):
                raise e
            else:
                # print 'PB', self.eval_ref(RangeCore.apply_one('add',RangeCore.apply_one('add',self.eval_ref("Cashflow!L33"),self.eval_ref("Cashflow!L35"),(39, 'L')),self.eval_ref("Cashflow!L37"),(39, 'L')))
                raise Exception("Problem evalling: %s for %s, %s" % (e,cell.address(),cell.python_expression)) 

        return cell.value


class ASTNode(object):
    """A generic node in the AST"""
    
    def __init__(self,token):
        super(ASTNode,self).__init__()
        self.token = token
    def __str__(self):
        return self.token.tvalue
    def __getattr__(self,name):
        return getattr(self.token,name)

    def children(self,ast):
        args = ast.predecessors(self)
        args = sorted(args,key=lambda x: ast.node[x]['pos'])
        return args

    def parent(self,ast):
        args = ast.successors(self)
        return args[0] if args else None

    def find_special_function(self, ast):
        found = False
        current = self

        special_functions = ['sumproduct', 'match']
        break_functions = ['index']

        while current is not None:
            if current.tvalue.lower() in special_functions:
                found = True
                break
            elif current.tvalue.lower() in break_functions:
                break
            else:
                current = current.parent(ast)

        return found

    def has_operator_or_func_parent(self, ast):
        found = False
        current = self

        while current is not None:
            if (current.ttype[:8] == 'operator' or current.ttype == 'function') and current.tvalue.lower() != 'if':
                found = True
                break
            else:
                current = current.parent(ast)

        return found

    def emit(self,ast,context=None):
        """Emit code"""
        self.token.tvalue
    
class OperatorNode(ASTNode):
    def __init__(self, args, ref):
        super(OperatorNode,self).__init__(args)
        self.ref = ref

        # convert the operator to python equivalents
        self.opmap = {
                 "^":"**",
                 "=":"==",
                 "&":"+",
                 "":"+" #union
                 }

        self.op_range_translator = {
            "*": "multiply",
            "/": "divide",
            "+": "add",
            "-": "substract",
            "==": "is_equal",
            "<>": "is_not_equal",
            ">": "is_strictly_superior",
            "<": "is_strictly_inferior",
            ">=": "is_superior_or_equal",
            "<=": "is_inferior_or_equal"
        }

    def emit(self,ast,context=None):
        xop = self.tvalue
        
        # Get the arguments
        args = self.children(ast)
        
        op = self.opmap.get(xop,xop)
        
        parent = self.parent(ast)
        # convert ":" operator to a range function
        if op == ":":
            # OFFSET HANDLER, when the first argument of OFFSET is a range i.e "A1:A2"
            if (parent is not None and
            (parent.tvalue == 'OFFSET' and 
             parent.children(ast)[0] == self)):
                return '"%s"' % ':'.join([a.emit(ast,context=context).replace('"', '') for a in args])
            else:
                return "self.eval_ref(%s)" % ','.join([a.emit(ast,context=context) for a in args])

         
        if self.ttype == "operator-prefix":
            return "RangeCore.apply_one('minus', %s, None, %s)" % (args[0].emit(ast,context=context), str(self.ref))

        if op in ["+", "-", "*", "/", "==", "<>", ">", "<", ">=", "<="]:
            is_special = self.find_special_function(ast)
            call = 'apply' + ('_all' if is_special else '_one')
            function = self.op_range_translator.get(op)

            arg1 = args[0]
            arg2 = args[1]

            return "RangeCore." + call + "(%s)" % ','.join(["'"+function+"'", str(arg1.emit(ast,context=context)), str(arg2.emit(ast,context=context)), str(self.ref)])

        parent = self.parent(ast)

        #TODO silly hack to work around the fact that None < 0 is True (happens on blank cells)
        if op == "<" or op == "<=":
            aa = args[0].emit(ast,context=context)
            ss = "(" + aa + " if " + aa + " is not None else float('inf'))" + op + args[1].emit(ast,context=context)
        elif op == ">" or op == ">=":
            aa = args[1].emit(ast,context=context)
            ss =  args[0].emit(ast,context=context) + op + "(" + aa + " if " + aa + " is not None else float('inf'))"
        else:
            ss = args[0].emit(ast,context=context) + op + args[1].emit(ast,context=context)
                    

        #avoid needless parentheses
        if parent and not isinstance(parent,FunctionNode):
            ss = "("+ ss + ")"          

        return ss

class OperandNode(ASTNode):
    def __init__(self,*args):
        super(OperandNode,self).__init__(*args)
    def emit(self,ast,context=None):
        t = self.tsubtype
        
        if t == "logical":
            return str(self.tvalue.lower() == "true")
        elif t == "text" or t == "error":
            #if the string contains quotes, escape them
            val = self.tvalue.replace('"','\\"')
            return '"' + val + '"'
        else:
            return str(self.tvalue)

class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range, e.g., A5, B3:C20 or INPUT """
    def __init__(self,args, ref):
        super(RangeNode,self).__init__(args)
        self.ref = ref # ref is the address of the reference cell  
    
    def get_cells(self):
        return resolve_range(self.tvalue)[0]
    
    def emit(self,ast,context=None):
        if isinstance(self.tvalue, ExcelError):
            print 'Excel Error Code found', self.tvalue
            return self.tvalue

        is_a_range = False
        is_a_named_range = self.tsubtype == "named_range"

        has_operator_or_func_parent = self.has_operator_or_func_parent(ast)

        if is_a_named_range:
            my_str = "'" + str(self) + "'" 
        else:
            # print 'Parsing a range into cells', self
            rng = self.tvalue.replace('$','')
            sheet = context + "!" if context else ""

            is_a_range = is_range(rng)

            if is_a_range:
                sh,start,end = split_range(rng)
            else:
                try:
                    sh,col,row = split_address(rng)
                except:
                    print 'WARNING: Unknown address: %s is not a cell/range reference, nor a named range' % str(rng)
                    sh = None

            if sh:
                my_str = '"' + rng + '"'
            else:
                my_str = '"' + sheet + rng + '"'

        to_eval = True
        # exception for formulas which use the address and not it content as ":" or "OFFSET"
        parent = self.parent(ast)
        # for OFFSET, it will also depends on the position in the formula (1st position required)
        if (parent is not None and
            (parent.tvalue == ':' or
            (parent.tvalue == 'OFFSET' and 
             parent.children(ast)[0] == self))):
            to_eval = False

        if parent is None and is_a_named_range: # When a named range is referenced in a cell without any prior operation
            return 'RangeCore.find_associated_values(' + str(self.ref) + ', self.eval_ref(' + my_str + '))[0]'
                        
        if to_eval == False:
            return my_str

        # OFFSET HANDLER
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            return 'RangeCore.find_associated_values(' + str(self.ref) + ', self.eval_ref(' + my_str + '))[0]'
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            return 'RangeCore.find_associated_values(' + str(self.ref) + ', self.eval_ref(' + my_str + '))[0]'

        # INDEX HANDLER
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[0] == self):


            if is_a_named_range:
                return 'resolve_range(self.named_ranges[' + my_str + '])'
            else:
                return 'resolve_range(' + my_str + ')'
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            return 'RangeCore.find_associated_values(' + str(self.ref) + ', self.eval_ref(' + my_str + '))[0]'
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            return 'RangeCore.find_associated_values(' + str(self.ref) + ', self.eval_ref(' + my_str + '))[0]'
        # elif is_a_range:
        #     return 'eval_range(' + str + ')'
        else:
            if (is_a_named_range or is_a_range) and not has_operator_or_func_parent:
                return 'RangeCore.find_associated_values(' + str(self.ref) + ', self.eval_ref(' + my_str + '))[0]'
            else:
                return 'self.eval_ref(' + my_str + ')'

        return my_str
    
class FunctionNode(ASTNode):
    """AST node representing a function call"""
    def __init__(self,*args):
        super(FunctionNode,self).__init__(*args)
        self.numargs = 0

        # map  excel functions onto their python equivalents
        self.funmap = excelfun.FUNCTION_MAP
        
    def emit(self,ast,context=None):
        fun = self.tvalue.lower()
        str = ''

        # Get the arguments
        args = self.children(ast)
        
        if fun == "atan2":
            # swap arguments
            str = "atan2(%s,%s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context))
        elif fun == "pi":
            # constant, no parens
            str = "pi"
        elif fun == "if":
            # inline the if
            if len(args) == 2:
                str = "%s if %s else 0" %(args[1].emit(ast,context=context),args[0].emit(ast,context=context))
            elif len(args) == 3:
                str = "(%s if %s else %s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context),args[2].emit(ast,context=context))
            else:
                raise Exception("if with %s arguments not supported" % len(args))

        elif fun == "array":
            str += '['
            if len(args) == 1:
                # only one row
                str += args[0].emit(ast,context=context)
            else:
                # multiple rows
                str += ",".join(['[' + n.emit(ast,context=context) + ']' for n in args])
                     
            str += ']'
        elif fun == "arrayrow":
            #simply create a list
            str += ",".join([n.emit(ast,context=context) for n in args])

        elif fun == "and":
            str = "all([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        elif fun == "or":
            str = "any([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        elif fun == "index":
            if self.parent(ast) is not None and self.parent(ast).tvalue == ':':
                str = 'index(' + ",".join([n.emit(ast,context=context) for n in args]) + ")"
            else:
                str = 'self.eval_ref(index(' + ",".join([n.emit(ast,context=context) for n in args]) + "))"
        elif fun == "offset":
            if self.parent(ast) is None or self.parent(ast).tvalue == ':':
                str = 'offset(' + ",".join([n.emit(ast,context=context) for n in args]) + ")"
            else:
                str = 'self.eval_ref(offset(' + ",".join([n.emit(ast,context=context) for n in args]) + "))"
        else:
            # map to the correct name
            f = self.funmap.get(fun,fun)
            str = f + "(" + ",".join([n.emit(ast,context=context) for n in args]) + ")"

        return str

def create_node(t, ref):
    """Simple factory function"""
    if t.ttype == "operand":
        if t.tsubtype == "range" or t.tsubtype == "named_range":
            return RangeNode(t, ref)
        else:
            return OperandNode(t)
    elif t.ttype == "function":
        return FunctionNode(t)
    elif t.ttype.startswith("operator"):
        return OperatorNode(t, ref)
    else:
        return ASTNode(t)

class Operator:
    """Small wrapper class to manage operators during shunting yard"""
    def __init__(self,value,precedence,associativity):
        self.value = value
        self.precedence = precedence
        self.associativity = associativity

def shunting_yard(expression, named_ranges, ref = '', tokenize_range = False):
    """
    Tokenize an excel formula expression into reverse polish notation
    
    Core algorithm taken from wikipedia with varargs extensions from
    http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-algorithm-to-allow-variable-numbers-of-arguments-to-functions/
    

    The ref is the cell address which is passed down to the actual compiled python code.
    Range basic operations signature require this reference, so it has to be written during OperatorNode.emit()
    https://github.com/iOiurson/koala/blob/master/koala/ast/graph.py#L292.

    This is needed because Excel range basic operations (+, -, * ...) are applied on matching cells.

    Example:
    Cell C2 has the following formula 'A1:A3 + B1:B3'.
    The output will actually be A2 + B2, because the formula is relative to cell C2.
    """

    #remove leading =
    if expression.startswith('='):
        expression = expression[1:]
        
    p = ExcelParser(tokenize_range = tokenize_range);
    p.parse(expression)

    # insert tokens for '(' and ')', to make things clearer below
    tokens = []
    for t in p.tokens.items:
        if t.ttype == "function" and t.tsubtype == "start":
            t.tsubtype = ""
            tokens.append(t)
            tokens.append(f_token('(','arglist','start'))
        elif t.ttype == "function" and t.tsubtype == "stop":
            tokens.append(f_token(')','arglist','stop'))
        elif t.ttype == "subexpression" and t.tsubtype == "start":
            t.tvalue = '('
            tokens.append(t)
        elif t.ttype == "subexpression" and t.tsubtype == "stop":
            t.tvalue = ')'
            tokens.append(t)
        elif t.ttype == "operand" and t.tsubtype == "range" and t.tvalue in named_ranges:
            t.tsubtype = "named_range"
            tokens.append(t)
        else:
            tokens.append(t)

    # print "==> ", "".join([t.tvalue for t in tokens]) 


    #http://office.microsoft.com/en-us/excel-help/calculation-operators-and-precedence-HP010078886.aspx
    operators = {}
    operators[':'] = Operator(':',8,'left')
    operators[''] = Operator(' ',8,'left')
    operators[','] = Operator(',',8,'left')
    operators['u-'] = Operator('u-',7,'left') #unary negation
    operators['%'] = Operator('%',6,'left')
    operators['^'] = Operator('^',5,'left')
    operators['*'] = Operator('*',4,'left')
    operators['/'] = Operator('/',4,'left')
    operators['+'] = Operator('+',3,'left')
    operators['-'] = Operator('-',3,'left')
    operators['&'] = Operator('&',2,'left')
    operators['='] = Operator('=',1,'left')
    operators['<'] = Operator('<',1,'left')
    operators['>'] = Operator('>',1,'left')
    operators['<='] = Operator('<=',1,'left')
    operators['>='] = Operator('>=',1,'left')
    operators['<>'] = Operator('<>',1,'left')
            
    output = collections.deque()
    stack = []
    were_values = []
    arg_count = []
    
    for t in tokens:
        if t.ttype == "operand":
            output.append(create_node(t, ref))
            if were_values:
                were_values.pop()
                were_values.append(True)
                
        elif t.ttype == "function":
            stack.append(t)
            arg_count.append(0)
            if were_values:
                were_values.pop()
                were_values.append(True)
            were_values.append(False)
            
        elif t.ttype == "argument":

            while stack and (stack[-1].tsubtype != "start"):
                output.append(create_node(stack.pop(), ref))   
            
            if were_values.pop(): arg_count[-1] += 1
            were_values.append(False)
            
            if not len(stack):
                raise Exception("Mismatched or misplaced parentheses")
        
        elif t.ttype.startswith('operator'):

            if t.ttype.endswith('-prefix') and t.tvalue =="-":
                o1 = operators['u-']
            else:
                o1 = operators[t.tvalue]

            while stack and stack[-1].ttype.startswith('operator'):
                
                if stack[-1].ttype.endswith('-prefix') and stack[-1].tvalue =="-":
                    o2 = operators['u-']
                else:
                    o2 = operators[stack[-1].tvalue]
                
                if ( (o1.associativity == "left" and o1.precedence <= o2.precedence)
                        or
                      (o1.associativity == "right" and o1.precedence < o2.precedence) ):
                    
                    output.append(create_node(stack.pop(), ref))
                else:
                    break
                
            stack.append(t)
        
        elif t.tsubtype == "start":
            stack.append(t)
            
        elif t.tsubtype == "stop":

            while stack and stack[-1].tsubtype != "start":
                output.append(create_node(stack.pop(), ref))
            
            if not stack:
                raise Exception("Mismatched or misplaced parentheses")
            
            stack.pop()

            if stack and stack[-1].ttype == "function":
                f = create_node(stack.pop(), ref)
                a = arg_count.pop()
                w = were_values.pop()
                if w: a += 1
                f.num_args = a
                #print f, "has ",a," args"
                output.append(f)

    while stack:
        if stack[-1].tsubtype == "start" or stack[-1].tsubtype == "stop":
            raise Exception("Mismatched or misplaced parentheses")
        
        output.append(create_node(stack.pop(), ref))

    #print "Stack is: ", "|".join(stack)
    #print "Output is: ", "|".join([x.tvalue for x in output])
    
    # convert to list
    return [x for x in output]
   
def build_ast(expression):
    """build an AST from an Excel formula expression in reverse polish notation"""
    #use a directed graph to store the tree
    G = DiGraph()
    
    stack = []
    
    for n in expression:
        # Since the graph does not maintain the order of adding nodes/edges
        # add an extra attribute 'pos' so we can always sort to the correct order
        if isinstance(n,OperatorNode):
            if n.ttype == "operator-infix":
                arg2 = stack.pop()
                arg1 = stack.pop()
                # Hack to write the name of sheet in 2argument address
                if(n.tvalue == ':'):
                    if '!' in arg1.tvalue and arg2.ttype == 'operand' and '!' not in arg2.tvalue:
                        arg2.tvalue = arg1.tvalue.split('!')[0] + '!' + arg2.tvalue
                    
                G.add_node(arg1,{'pos':1})
                G.add_node(arg2,{'pos':2})
                G.add_edge(arg1, n)
                G.add_edge(arg2, n)
            else:
                arg1 = stack.pop()
                G.add_node(arg1,{'pos':1})
                G.add_edge(arg1, n)
                
        elif isinstance(n,FunctionNode):
            args = []
            for _ in range(n.num_args):
                try:
                    args.append(stack.pop())
                except:
                    print 'STACK', stack, n, expression
                    raise Exception()
            #try:
                # args = [stack.pop() for _ in range(n.num_args)]
            #except:
            #        print 'STACK', stack, type(n)
            #        raise Exception('prut')
            args.reverse()
            for i,a in enumerate(args):
                G.add_node(a,{'pos':i})
                G.add_edge(a,n)

        else:
            G.add_node(n,{'pos':0})

        stack.append(n)

    return G,stack.pop()

def find_node(G, seed_address):
    for i,seed in enumerate(G.nodes()):
        if seed.address() == seed_address:
            return seed

def make_subgraph(G, seed, direction = "ascending"):
    subgraph = networkx.DiGraph()
    if direction == "ascending":
        todo = map(lambda n: (seed,n), G.predecessors(seed))
    else:
        todo = map(lambda n: (seed,n), G.successors(seed))
    while len(todo) > 0:
        neighbor, current = todo.pop()
        subgraph.add_node(current)
        subgraph.add_edge(neighbor, current)
        if direction == "ascending":
            nexts = G.predecessors(current)
        else:
            nexts = G.successors(current)
        for n in nexts:            
            if n not in subgraph.nodes():
                todo += [(current,n)]

    return subgraph

def cell2code(named_ranges, cell, sheet):
    """Generate python code for the given cell"""
    if cell.formula:
        ref = parse_cell_address(cell.address()) if not cell.is_named_range else None
        e = shunting_yard(cell.formula or str(cell.value), named_ranges, ref=ref)
        ast,root = build_ast(e)
        code = root.emit(ast, context=sheet)
    else:
        ast = None
        code = str('"' + cell.value + '"' if isinstance(cell.value,unicode) else cell.value)
    return code,ast

class ExcelCompiler(object):
    """Class responsible for taking cells and named_range and create a graph
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, file, ignore_sheets = []):
        print "___### Initializing Excel Compiler ###___"

        file_name = os.path.abspath(file)
        # Decompose subfiles structure in zip file
        archive = read_archive(file_name)
        # Parse cells
        self.cells = read_cells(archive, ignore_sheets)
        # Parse named_range { name (ExampleName) -> address (Sheet!A1:A10)}
        self.named_ranges = read_named_ranges(archive)
        self.Range = RangeFactory(self.cells)


    def clean_volatile(self):
        print '___### Cleaning volatiles ###___'

        sp = Spreadsheet(networkx.DiGraph(),self.cells, self.named_ranges)

        cleaned_cells, cleaned_ranged_names = sp.clean_volatile()
        self.cells = cleaned_cells

        self.named_ranges = cleaned_ranged_names

    
            
    def gen_graph(self, outputs = None):
        print '___### Generating Graph ###___'

        if outputs is None:
            seeds = list(flatten(self.cells.values()))
        else:
            seeds = []
            for o in outputs:
                if o in self.named_ranges:
                    reference = self.named_ranges[o]
                    if is_range(reference):

                        rng = self.Range(reference)
                        virtual_cell = Cell(o, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                        seeds.append(virtual_cell)
                    else:
                        virtual_cell = Cell(o, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True)
                        seeds.append(virtual_cell)
                else:
                    if is_range(o):
                        raise Exception("Your want a output range ?")
                    else:
                        seeds.append(self.cells[o])


        # print "Seeds %s cells" % len(seeds)
        
        # cells to analyze: only formulas
        todo = [s for s in seeds if s.formula]
        steps = [i for i,s in enumerate(todo)]

        # print "%s cells on the todo list" % len(todo)

        # map of all cells
        cellmap = dict([(x.address(),x) for x in seeds])
    
        # directed graph
        G = nx.DiGraph()

        # match the info in cellmap
        for c in cellmap.itervalues(): G.add_node(c)

        while todo:
            c1 = todo.pop()
            step = steps.pop()
            cursheet = c1.sheet

            ###### 1) looking for cell c1 dependencies ####################

            # in case a formula, get all cells that are arguments
            pystr, ast = cell2code(self.named_ranges, c1, cursheet)
            # set the code & compile it (will flag problems sooner rather than later)
            c1.python_expression = pystr
            c1.compile()    
            
            # get all the cells/ranges this formula refers to
            deps = [x.tvalue.replace('$','') for x in ast.nodes() if isinstance(x,RangeNode)]
            # remove dupes
            deps = uniqueify(deps)


            ###### 2) connect dependencies in cells in graph ####################

            # ### LOG
            # tmp = []
            # for dep in deps:
            #     if dep not in self.named_ranges:
            #         if "!" not in dep and cursheet != None:
            #             dep = cursheet + "!" + dep
            #     if dep not in cellmap:
            #         tmp.append(dep)
            # #deps = tmp
            # logStep = "%s %s = %s " % ('|'*step, c1.address(), '',)
            # print logStep

            # if len(deps) > 1 and 'L' in deps[0] and deps[0] == deps[-1].replace('DG','L'):
            #     print logStep, "[%s...%s]" % (deps[0], deps[-1])
            # elif len(deps) > 0:
            #     print logStep, "->", deps
            # else:
            #     print logStep, "done"
            
            for dep in deps:
                # we need an absolute address
                if dep not in self.named_ranges and "!" not in dep and cursheet != None:
                    dep = cursheet + "!" + dep

                # Named_ranges + ranges already parsed (previous iterations)
                if dep in cellmap:
                    origins = [cellmap[dep]]
                    target = cellmap[c1.address()]
                # if the dependency is a multi-cell range, create a range object
                elif is_range(dep) or (dep in self.named_ranges and is_range(self.named_ranges[dep])):

                    if dep in self.named_ranges:
                        reference = self.named_ranges[dep]
                    else:
                        reference = dep

                    rng = self.Range(reference)

                    formulas_in_dep = []
                    for c in rng.cells:
                        if c in self.cells:
                            formulas_in_dep.append(self.cells[c].formula)
                        else:
                            # raise Exception( '%s unavailable' % c)
                            formulas_in_dep.append(None)
                
                    virtual_cell = Cell(dep, None, value = rng, formula = reference, is_range = True, is_named_range = True )

                    # save the range
                    cellmap[dep] = virtual_cell
                    # add an edge from the range to the parent
                    G.add_node(virtual_cell)
                    # Cell(A1:A10) -> c1 or Cell(ExampleName) -> c1
                    G.add_edge(virtual_cell, cellmap[c1.address()])
                    # cells in the range should point to the range as their parent
                    target = virtual_cell 
                    origins = []
                    for (child, value, formula) in zip(rng.cells, rng.value, formulas_in_dep):
                        if child not in cellmap:
                            cell_is_range = isinstance(value, RangeCore)
                            origins.append(Cell(child, None, value = value, formula = formula, is_range = cell_is_range, is_named_range = False))  
                        else:
                            origins.append(cellmap[child])   
                else:
                    # not a range 
                    if dep in self.named_ranges:
                        reference = self.named_ranges[dep]
                    else:
                        reference = dep


                    if reference in self.cells:
                        if dep in self.named_ranges:
                            virtual_cell = Cell(dep, None, value = self.cells[reference].value, formula = reference, is_range = False, is_named_range = True )
                            origins = [virtual_cell]
                        else:
                            origins = [self.cells[reference]] 
                    else:
                        virtual_cell = Cell(dep, None, value = None, formula = None, is_range = False, is_named_range = True )
                        origins = [virtual_cell]

                    target = cellmap[c1.address()]


                # process each cell                    
                for c2 in flatten(origins):
                    
                    # if we havent treated this cell allready
                    if c2.address() not in cellmap:
                        if c2.formula:
                            # cell with a formula, needs to be added to the todo list
                            todo.append(c2)
                            steps.append(step+1)
                        else:
                            # constant cell, no need for further processing, just remember to set the code
                            pystr,ast = cell2code(self.named_ranges, c2, cursheet)
                            c2.python_expression = pystr
                            c2.compile()     
                        
                        # save in the cellmap
                        cellmap[c2.address()] = c2
                        # add to the graph
                        G.add_node(c2)
                        
                    # add an edge from the cell to the parent (range or cell)
                    if(target != []):
                        # print "Adding edge %s --> %s" % (c2.address(), target.address())
                        G.add_edge(cellmap[c2.address()],target)

        print "Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))
        undirected = networkx.Graph(G)
        # print "Number of connected components %s", str(number_connected_components(undirected))

        return Spreadsheet(G, cellmap, self.named_ranges, outputs = outputs)




