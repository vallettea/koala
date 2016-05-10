


import os.path
import textwrap
import koala.ast.excellib as excelfun
from koala.ast.excellib import *
from koala.ast.excelutils import *
from math import *
from collections import OrderedDict

import networkx
from networkx.classes.digraph import DiGraph
from networkx.drawing.nx_pydot import write_dot
from networkx.drawing.nx_pylab import draw, draw_circular
from networkx.readwrite.gexf import write_gexf
from networkx.readwrite import json_graph
from networkx.algorithms import number_connected_components

import networkx as nx

from astutils import find_node, subgraph

from tokenizer import ExcelParser, f_token, shunting_yard
import cPickle
import logging
from itertools import chain

from Range import Range, find_associated_values, parse_cell_address, get_cell_address
from OffsetParser import OffsetParser

import json
import gzip

from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from ..excel.utils import rows_from_range

class Spreadsheet(object):
    def __init__(self, G, cellmap, named_ranges, ranges):
        super(Spreadsheet,self).__init__()
        self.G = G
        self.cellmap = cellmap
        self.named_ranges = named_ranges
        self.ranges = ranges
        self.params = None
        self.history = dict()
        self.count = 0


    def dump(self, fname):
        data = json_graph.node_link_data(self.G)
        # save nodes as simple objects
        nodes = []
        for node in data["nodes"]:
            cell = node["id"]
            if isinstance(cell.value, OrderedDict):
                value = zip(cell.value.keys(), cell.value.values())
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
        # save ranges as simple objects
        ranges = {}
        for k,r in self.ranges.items():
            ranges[k] = zip(r.keys(), r.values())
        data["ranges"] = ranges
        data["named_ranges"] = self.named_ranges
        with gzip.GzipFile(fname, 'w') as outfile:
            outfile.write(json.dumps(data))

    @staticmethod
    def load(fname):
        with gzip.GzipFile(fname, 'r') as infile:
            data = json.loads(infile.read())
        def cell_from_dict(d):
            if hasattr(d["value"], '__iter__'):
                value = OrderedDict(map(lambda x: (tuple(x[0]), x[1]), d["value"]))
            else:
                value = d["value"]
            return {"id": Cell(d["address"], None, value=value, formula=d["formula"], is_named_range=d["is_named_range"], always_eval=d["always_eval"])}
        nodes = map(cell_from_dict, data["nodes"])
        data["nodes"] = nodes
        G = json_graph.node_link_graph(data)
        ranges = {k: OrderedDict(map(lambda x: (tuple(x[0]), x[1]), v)) for k,v in data["ranges"].items()}
        return Spreadsheet(G, G.nodes(), data["named_ranges"], ranges)


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
    
    def set_value(self,cell,val,is_addr=True):
        if is_addr:
            address = cell.replace('$','')
            # for s in self.cellmap:
            #     print 'c', self.cellmap[s].address(), self.cellmap[s].value
            cell = self.cellmap[address]

        if cell.is_named_range:
            # Take care of the case where named_range is not directly a cell address (type offset ...)
            # It will raise an exception, but we want this to prevent wrong usage
            return self.set_value(self.cellmap[cell.formula], val,False)

        if cell.value != val:
            # reset the node + its dependencies
            self.reset(cell)
            # set the value
            cell.value = val

    def reset(self, cell):
        addr = cell.address()
        if cell.value is None and addr not in self.named_ranges: return

        # update depending ranges
        if addr in self.ranges:
            self.ranges[addr].reset()

        cell.value = None
        map(self.reset,self.G.successors_iter(cell))

    def print_value_tree(self,addr,indent):
        cell = self.cellmap[addr]
        print "%s %s = %s" % (" "*indent,addr,cell.value)
        for c in self.G.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    def recalculate(self):
        for c in self.cellmap.values():
            if isinstance(c,CellRange):
                self.evaluate_range(c,is_addr=False)
            else:
                self.evaluate(c,is_addr=False)
                
    def evaluate_range(self,rng,is_addr=True):

        if is_addr:
            rng = self.cellmap[rng]

        # its important that [] gets treated ad false here
        if rng.value:
            return rng.value

        cells,nrows,ncols = rng.celladdrs,rng.nrows,rng.ncols

        cells = list(flatten(cells))

        values = [ self.evaluate(c) for c in cells ]

        data = Range(cells, values)
        rng.value = data
        
        return data

    def evaluate(self,cell,is_addr=True):

        if is_addr:
            try:
                # print '->', cell
                cell = self.cellmap[cell]

            except:
                # print 'Empty cell at '+ cell
                return None

        # no formula, fixed value
        if not cell.formula or not cell.always_eval and cell.value != None:
            #print "returning constant or cached value for ", cell.address()
            return cell.value
        
        def update_range(range):
            # print 'update range', range
            for key, value in range.items():
                if value is None:
                    addr = get_cell_address(range.sheet, key)
                    range[key] = self.evaluate(addr)

            # print 'updated range', range
            return range
        # recalculate formula
        # the compiled expression calls this function
        def eval_ref(addr1, addr2 = None):
            if addr2 == None:
                if addr1 in self.ranges:
                    range1 = self.ranges[addr1]
                    return update_range(range1)

                elif addr1 in self.named_ranges:
                    return self.evaluate(addr1)
                elif not is_range(addr1): # addr1 = Sheet1!A1 or A1, maybe this may never happen
                    # print 'REF1 is not a range'
                    return self.evaluate(addr1)
                else: # addr1 = Sheet1!A1:A2 or Sheet1!A1:Sheet1!A2
                    addr1, addr2 = addr1.split(':')
                    if '!' in addr1:
                        sheet = addr1.split('!')[0]
                    else:
                        sheet = None
                    if '!' in addr2:
                        addr2 = addr2.split('!')[1]
                    # print 'REF1 is a range'
                    return self.evaluate_range(CellRange('%s:%s' % (addr1, addr2),sheet), False)
            else:  # addr1 = Sheet1!A1, addr2 = Sheet1!A2
                # print 'REF2 is not none'
                if '!' in addr1:
                    sheet = addr1.split('!')[0]
                else:
                    sheet = None
                if '!' in addr2:
                    addr2 = addr2.split('!')[1]
                return self.evaluate_range(CellRange('%s:%s' % (addr1, addr2),sheet), False)

        try:
            # print "Evalling: %s, %s" % (cell.address(),cell.python_expression)
            vv = eval(cell.compiled_expression)
            # if vv is None:
            #     print "WARNING %s is None" % (cell.address())
            # elif isinstance(vv, (List, list)):
            #     print 'Output is list => converting', cell.index
            #     vv = vv[cell.index]
            cell.value = vv

            # DEBUG: saving differences
            if cell.address() in self.history:
                ori_value = self.history[cell.address()]['original']
                if is_number(ori_value) and is_number(cell.value) and abs(float(ori_value) - float(cell.value)) > 0.001:
                    self.count += 1
                    self.history[cell.address()]['formula'] = str(cell.formula)
                    self.history[cell.address()]['priority'] = self.count
                    self.history[cell.address()]['python'] = str(cell.python_expression)

                self.history[cell.address()]['new'] = str(cell.value)
            else:
                self.history[cell.address()] = {'new': str(cell.value)}

        except Exception as e:
            if e.message.startswith("Problem evalling"):
                raise e
            else:
                raise Exception("Problem evalling: %s for %s, %s" % (e,cell.address(),cell.python_expression)) 

        try:
            return cell.value
        except:
            for f in missing_functions:
                print 'MISSING', f

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
        #args.reverse()
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
            # print 'VERIF', current.tvalue.lower()

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
            # print 'VERIF', current.tvalue.lower(), current.ttype

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
                return "eval_ref(%s)" % ','.join([a.emit(ast,context=context) for a in args])

         
        if self.ttype == "operator-prefix":
            return "Range.apply_one('minus', %s, None, %s)" % (args[0].emit(ast,context=context), str(self.ref))

        if op in ["+", "-", "*", "/", "==", "<>", ">", "<", ">=", "<="]:
            is_special = self.find_special_function(ast)
            call = 'apply' + ('_all' if is_special else '_one')
            function = self.op_range_translator.get(op)

            arg1 = args[0]
            arg2 = args[1]

            return "Range." + call + "(%s)" % ','.join(["'"+function+"'", str(arg1.emit(ast,context=context)), str(arg2.emit(ast,context=context)), str(self.ref)])

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
        is_a_range = False
        is_a_named_range = self.tsubtype == "named_range"

        has_operator_or_func_parent = self.has_operator_or_func_parent(ast)

        if is_a_named_range:
            # print 'RANGE', str(self)
            my_str = "'" + str(self) + "'" 
        else:
            # print 'Parsing a range into cells', self
            rng = self.tvalue.replace('$','')
            sheet = context + "!" if context else ""

            is_a_range = is_range(rng)

            if is_a_range:
                sh,start,end = split_range(rng)
            else:
                sh,col,row = split_address(rng)

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
            return 'find_associated_values(' + str(self.ref) + ', eval_ref(' + my_str + '))[0]'
                        
        if to_eval == False:
            return my_str

        # OFFSET HANDLER
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            return 'find_associated_values(' + str(self.ref) + ', eval_ref(' + my_str + '))[0]'
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            return 'find_associated_values(' + str(self.ref) + ', eval_ref(' + my_str + '))[0]'

        # INDEX HANDLER
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[0] == self):


            if is_a_named_range:
                return 'resolve_range(self.named_ranges[' + my_str + '])'
            else:
                return 'resolve_range(' + my_str + ')'
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            return 'find_associated_values(' + str(self.ref) + ', eval_ref(' + my_str + '))[0]'
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            return 'find_associated_values(' + str(self.ref) + ', eval_ref(' + my_str + '))[0]'
        # elif is_a_range:
        #     return 'eval_range(' + str + ')'
        else:
            if (is_a_named_range or is_a_range) and not has_operator_or_func_parent:
                return 'find_associated_values(' + str(self.ref) + ', eval_ref(' + my_str + '))[0]'
            else:
                return 'eval_ref(' + my_str + ')'

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
        elif fun == "index": # might not be necessary
            if self.parent(ast) is not None and self.parent(ast).tvalue == ':':
                str = 'index(' + ",".join([n.emit(ast,context=context) for n in args]) + ")"
            else:
                str = 'eval_ref(index(' + ",".join([n.emit(ast,context=context) for n in args]) + "))"
        elif fun == "offset":
            if self.parent(ast) is None or self.parent(ast).tvalue == ':':
                str = 'offset(' + ",".join([n.emit(ast,context=context) for n in args]) + ")"
            else:
                str = 'eval_ref(offset(' + ",".join([n.emit(ast,context=context) for n in args]) + "))"
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

def shunting_yard(expression, named_ranges, ref = ''):
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
        
    p = ExcelParser();
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
            args = [stack.pop() for _ in range(n.num_args)]
            args.reverse()
            for i,a in enumerate(args):
                G.add_node(a,{'pos':i})
                G.add_edge(a,n)
            #for i in range(n.num_args):
            #    G.add_edge(stack.pop(),n)
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

class ExcelCompiler(object):
    """Class responsible for taking cells and named_range and create a graph
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, file, ignore_sheets = [], parse_offsets = False):

        file_name = os.path.abspath(file)
        # Decompose subfiles structure in zip file
        archive = read_archive(file_name)
        # Parse cells
        self.cells = read_cells(archive, ignore_sheets)
        # Parse named_range
        self.named_ranges = read_named_ranges(archive)
        # Remove offsets
        if parse_offsets:
            parser = OffsetParser(self.cells, self.named_ranges)
            nb = 0
            for k,v in self.named_ranges.items():
                if 'OFFSET' in v:
                    self.named_ranges[k] = parser.parseOffsets(v)
                    nb +=1
            for k,cell in self.cells.items():
                if cell.formula and 'OFFSET' in cell.formula:
                    f = parser.parseOffsets(cell.formula)
                    c = Cell(cell.address(), cell.sheet, cell.value, f, cell.is_named_range, cell.always_eval)
                    self.cells[k] = c
                    nb +=1
            print "%s offsets removed" % str(nb)
        
        # Transform named_ranges in artificial ranges
        self.ranges = {}
        for n in self.named_ranges:
            reference = self.named_ranges[n]
            if is_range(reference):
                if 'OFFSET' not in reference:
                    range_cells, nrow, ncol = resolve_range(reference)

                    range_cells = list(flatten(range_cells))
                    range_values = []

                    for cell in range_cells:
                        if cell in self.cells: # this is to avoid Depreciation!A5 and other empty cells due to tR named range
                            range_values.append(self.cells[cell].value)
                        else:
                            range_values.append(None)

                    my_range = Range(range_cells, range_values)
                    self.ranges[n] = my_range
                    self.cells[n] = Cell(n, None, my_range, n, True )
                else:
                    self.cells[n] = Cell(n, None, None, self.named_ranges[n], True )
            else:
                if reference in self.cells:
                    self.cells[n] = Cell(n, None, self.cells[reference].value, reference, True )
                else:
                    self.cells[n] = Cell(n, None, None, reference, True )

    def cell2code(self, cell, sheet):
        """Generate python code for the given cell"""
        if cell.formula:
            ref = parse_cell_address(cell.address()) if not cell.is_named_range else None
            e = shunting_yard(cell.formula or str(cell.value), self.named_ranges, ref)
            ast,root = build_ast(e)
            code = root.emit(ast, context=sheet)
        else:
            ast = None
            code = str('"' + cell.value + '"' if isinstance(cell.value,unicode) else cell.value)
        return code,ast

    
            
    def gen_graph(self, outputs = None, inputs = None):
        
        if outputs is None:
            seeds = list(flatten(self.cells.values()))
        else:
            seeds = [self.cells[o] for o in outputs]

        print "Seeds %s cells" % len(seeds)
        # only keep seeds with formulas or numbers
        seeds = [s for s in seeds if s.formula or isinstance(s.value,(int, float, str))]

        print "%s filtered seeds " % len(seeds)
        
        # cells to analyze: only formulas
        todo = [s for s in seeds if s.formula]

        print "%s cells on the todo list" % len(todo)

        # map of all cells
        cellmap = dict([(x.address(),x) for x in seeds])
    
        # directed graph
        G = nx.DiGraph()

        # match the info in cellmap
        for c in cellmap.itervalues(): G.add_node(c)

        while todo:
            c1 = todo.pop()
            
            # print "============= Handling ", c1.address()
            cursheet = c1.sheet
            
            if c1.address() in self.ranges:
                deps = []
                for c in self.ranges[c1.address()].cells:
                    deps.append(c)
            else:
                # parse the formula into code
                pystr, ast = self.cell2code(c1, cursheet)
                # set the code & compile it (will flag problems sooner rather than later)
                c1.python_expression = pystr
                c1.compile()    
                
                # get all the cells/ranges this formula refers to
                deps = [x.tvalue.replace('$','') for x in ast.nodes() if isinstance(x,RangeNode)]
                # remove dupes
                deps = uniqueify(deps)

            for dep in deps:
                if dep in self.named_ranges:
                    cells = [self.cells[dep]]
                    target = cellmap[c1.address()]
                # if the dependency is a multi-cell range, create a range object
                elif is_range(dep):
                    # this will make sure we always have an absolute address
                    rng = CellRange(dep, sheet=cursheet)
                    
                    if rng.address() in cellmap:
                        # already dealt with this range
                        # add an edge from the range to the parent
                        G.add_edge(cellmap[rng.address()],cellmap[c1.address()])
                        continue
                    else:
                        # turn into cell objects
                        if "!" in dep:
                            sheet_name, ref = dep.split("!")
                        else:
                            sheet_name = cursheet
                            ref = dep
                        cells_refs = list(rows_from_range(ref))                       
                        cells = [self.cells[sheet_name +"!"+ ref] for ref in list(chain(*cells_refs)) if sheet_name +"!"+ ref in self.cells]

                        # get the values so we can set the range value
                        rng.value = [c.value for c in cells]
                        
                        # my_range = Range(cells, rng.value)
                        # self.ranges[dep] = my_range

                        # save the range
                        cellmap[rng.address()] = rng
                        # add an edge from the range to the parent
                        G.add_node(rng)
                        G.add_edge(rng,cellmap[c1.address()])
                        # cells in the range should point to the range as their parent
                        target = rng
                else:
                    # not a range, create the cell object
                    if "!" in dep:
                        sheet_name, ref = dep.split("!")
                    else:
                        sheet_name = cursheet
                        ref = dep
                    try:
                        temp = self.cells[ref] if ref in self.named_ranges else self.cells[sheet_name +"!"+ ref]
                        cells = [temp]
                        target = cellmap[c1.address()]
                    except:
                        cells = []
                        target = []

                # process each cell                    
                for c2 in flatten(cells):
                    # if we havent treated this cell allready
                    if c2.address() not in cellmap:
                        if c2.formula:
                            # cell with a formula, needs to be added to the todo list
                            todo.append(c2)
                        else:
                            # constant cell, no need for further processing, just remember to set the code
                            pystr,ast = self.cell2code(c2, cursheet)
                            c2.python_expression = pystr
                            c2.compile()     
                        
                        # save in the cellmap
                        cellmap[c2.address()] = c2
                        # add to the graph
                        G.add_node(c2)
                        
                    # add an edge from the cell to the parent (range or cell)
                    if(target != []):
                        G.add_edge(cellmap[c2.address()],target)

        print "Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))
        undirected = networkx.Graph(G)
        print "Number of connected components %s", str(number_connected_components(undirected))
        

        if inputs != None:
            # get all the cells impacted by inputs
            dependencies = set()
            for input_address in inputs:
                seed = find_node(G, input_address)
                if seed != None:
                    g = make_subgraph(G, seed, "descending")
                    dependencies = dependencies.union(g.nodes())
                else:
                    print "Node corresponding to %s not in graph" % input_address
            print "%s cells depending on inputs" % str(len(dependencies))
            # print map(lambda x: x.address(), dependencies)

            # prune the graph and set all cell independent of input to const
            subgraph = networkx.DiGraph()
            new_cellmap = {}
            for output_address in outputs:
                seed = find_node(G, output_address)
                todo = map(lambda n: (seed,n), G.predecessors(seed))

                while len(todo) > 0:
                    current, pred = todo.pop()
                    # print "==========================="
                    # print current.address(), pred.address()
                    if current in dependencies:
                        if pred in dependencies:
                            subgraph.add_edge(pred, current)
                            new_cellmap[pred.address()] = pred
                            new_cellmap[current.address()] = current

                            nexts = G.predecessors(pred)
                            for n in nexts:            
                                if n not in subgraph.nodes():
                                    todo += [(pred,n)]
                        else:
                            if pred.address() not in new_cellmap:
                                const_node = Cell(pred.address(), pred.sheet, value=pred.value, formula=None, is_named_range=pred.is_named_range, always_eval=pred.always_eval)
                                pystr,ast = self.cell2code(const_node, pred.sheet)
                                const_node.python_expression = pystr
                                const_node.compile()     
                            else:
                                const_node = new_cellmap[pred.address()]
                            subgraph.add_edge(const_node, current)
                            new_cellmap[const_node.address()] = const_node

                    else:
                        if pred.address() not in new_cellmap:
                            const_node = Cell(pred.address(), pred.sheet, value=pred.value, formula=None, is_named_range=pred.is_named_range, always_eval=pred.always_eval)
                            pystr,ast = self.cell2code(const_node, pred.sheet)
                            const_node.python_expression = pystr
                            const_node.compile()     
                        else:
                            const_node = new_cellmap[pred.address()]
                        subgraph.add_node(const_node)
                        new_cellmap[const_node.address()] = const_node
    
                        

            G = subgraph
            cellmap = new_cellmap
            print "Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))
            undirected = networkx.Graph(G)
            print "Number of connected components %s", str(number_connected_components(undirected))
            # print map(lambda x: x.address(), G.nodes())

        sp = Spreadsheet(G,cellmap, self.named_ranges, self.ranges)
        
        return sp

