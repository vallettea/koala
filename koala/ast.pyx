# cython: profile=True

import networkx
import collections

import os.path
# from math import *

import networkx
from networkx.classes.digraph import DiGraph

# from excellib import *
from utils import uniqueify, flatten
from Cell import Cell
from Range import parse_cell_address
from tokenizer import ExcelParser, f_token, shunting_yard
from astnodes import *


def create_node(t, ref, debug = False):
    """Simple factory function"""
    if t.ttype == "operand":
        if t.tsubtype == "range" or t.tsubtype == "named_range":
            return RangeNode(t, ref, debug = debug)
        else:
            return OperandNode(t)
    elif t.ttype == "function":
        return FunctionNode(t, ref, debug = debug)
    elif t.ttype.startswith("operator"):
        return OperatorNode(t, ref, debug = debug)
    else:
        return ASTNode(t, debug = debug)

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

# Whats the difference between subgraph() and make_subgraph() ?

def subgraph(G, seed):
    subgraph = networkx.DiGraph()
    todo = map(lambda n: (seed,n), G.predecessors(seed))
    while len(todo) > 1:
        previous, current = todo.pop()
        addr = current.address()
        subgraph.add_node(current)
        subgraph.add_edge(previous, current)
        for n in G.predecessors(current):            
            if n not in subgraph.nodes():
                todo += [(current,n)]

    return subgraph

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
        code = str('"' + cell.value.encode('utf-8') + '"' if isinstance(cell.value,unicode) else cell.value)
    return code,ast



def graph_from_seeds(seeds, cell_source):
    """
    This creates/updates a networkx graph from a list of cells.

    The graph is created when the cell_source is an instance of ExcelCompiler
    The graph is updated when the cell_source is an instance of Spreadsheet
    """

    # when called from Spreadsheet instance, use the Spreadsheet cellmap and graph 
    if hasattr(cell_source, 'G'): # ~ cell_source is a Spreadsheet
        cellmap = cell_source.cellmap
        cells = cellmap
        G = cell_source.G
        for c in seeds: 
            G.add_node(c)
            cellmap[c.address()] = c
    # when called from ExcelCompiler instance, construct cellmap and graph from seeds 
    # elif isinstance(cell_source, ExcelCompiler):
    else: # ~ cell_source is a Spreadsheet
        cellmap = dict([(x.address(),x) for x in seeds])
        cells = cell_source.cells
        # directed graph
        G = networkx.DiGraph()
        # match the info in cellmap
        for c in cellmap.itervalues(): G.add_node(c)

    # cells to analyze: only formulas
    todo = [s for s in seeds if s.formula]
    steps = [i for i,s in enumerate(todo)]

    while todo:
        c1 = todo.pop()
        step = steps.pop()
        cursheet = c1.sheet

        ###### 1) looking for cell c1 dependencies ####################

        # in case a formula, get all cells that are arguments
        pystr, ast = cell2code(cell_source.named_ranges, c1, cursheet)
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
        #     if dep not in cell_source.named_ranges:
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
            # this is to avoid :A1 or A1: dep due to clean_volatiles() returning an ExcelError
            if dep.startswith(':') or dep.endswith(':'):
                dep = dep.replace(':', '')

            # we need an absolute address
            if dep not in cell_source.named_ranges and "!" not in dep and cursheet != None:
                dep = cursheet + "!" + dep

            # Named_ranges + ranges already parsed (previous iterations)
            if dep in cellmap:
                origins = [cellmap[dep]]
                target = cellmap[c1.address()]
            # if the dependency is a multi-cell range, create a range object
            elif is_range(dep) or (dep in cell_source.named_ranges and is_range(cell_source.named_ranges[dep])):

                if dep in cell_source.named_ranges:
                    reference = cell_source.named_ranges[dep]
                else:
                    reference = dep
                
                rng = cell_source.Range(reference)

                if len(rng.keys()) != 0: # could be better, but can't check on Exception types here...
                    formulas_in_dep = []
                    for c in rng.addresses:
                        if c in cells:
                            formulas_in_dep.append(cells[c].formula)
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

                if len(rng.keys()) != 0: # could be better, but can't check on Exception types here...
                    for child in rng.addresses:
                        if child not in cellmap:
                            origins.append(cells[child])  
                        else:
                            origins.append(cellmap[child])   
            else:
                # not a range 
                if dep in cell_source.named_ranges:
                    reference = cell_source.named_ranges[dep]
                else:
                    reference = dep


                if reference in cells:
                    if dep in cell_source.named_ranges:
                        virtual_cell = Cell(dep, None, value = cells[reference].value, formula = reference, is_range = False, is_named_range = True )
                        origins = [virtual_cell]
                    else:
                        origins = [cells[reference]] 
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
                        pystr,ast = cell2code(cell_source.named_ranges, c2, cursheet)
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

    return (cellmap, G)
