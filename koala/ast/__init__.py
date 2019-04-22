from __future__ import absolute_import
# cython: profile=True

import collections
import six

import networkx
from networkx.classes.digraph import DiGraph
from openpyxl.compat import unicode

from koala.utils import uniqueify, flatten, max_dimension, col2num, resolve_range
from koala.Cell import Cell
from koala.Range import parse_cell_address
from koala.tokenizer import ExcelParser, f_token
from .astnodes import *


def create_node(t, ref = None, debug = False):
    """Simple factory function"""
    if t.ttype == "operand":
        if t.tsubtype in ["range", "named_range", "pointer"] :
            # print 'Creating Node', t.tvalue, t.tsubtype
            return RangeNode(t, ref, debug = debug)
        else:
            return OperandNode(t)
    elif t.ttype == "function":
        return FunctionNode(t, ref, debug = debug)
    elif t.ttype.startswith("operator"):
        return OperatorNode(t, ref, debug = debug)
    else:
        return ASTNode(t, debug = debug)


class Operator(object):
    """Small wrapper class to manage operators during shunting yard"""
    def __init__(self,value,precedence,associativity):
        self.value = value
        self.precedence = precedence
        self.associativity = associativity


def shunting_yard(expression, named_ranges, ref = None, tokenize_range = False):
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

    new_tokens = []

    # reconstruct expressions with ':' and replace the corresponding tokens by the reconstructed expression
    if not tokenize_range:
        for index, token in enumerate(tokens):
            new_tokens.append(token)

            if type(token.tvalue) == str or type(token.tvalue) == unicode:

                if token.tvalue.startswith(':'): # example -> :OFFSET( or simply :A10
                    depth = 0
                    expr = ''

                    rev = reversed(tokens[:index])

                    for t in rev: # going backwards, 'stop' starts, 'start' stops
                        if t.tsubtype == 'stop':
                            depth += 1
                        elif depth > 0 and t.tsubtype == 'start':
                            depth -= 1

                        expr = t.tvalue + expr

                        new_tokens.pop()

                        if depth == 0:
                            new_tokens.pop() # these 2 lines are needed to remove INDEX()
                            new_tokens.pop()
                            expr = six.next(rev).tvalue + expr
                            break

                    expr += token.tvalue

                    depth = 0

                    if token.tvalue[1:] in ['OFFSET', 'INDEX']:
                        for t in tokens[(index + 1):]:
                            if t.tsubtype == 'start':
                                depth += 1
                            elif depth > 0 and t.tsubtype == 'stop':
                                depth -= 1

                            expr += t.tvalue

                            tokens.remove(t)

                            if depth == 0:
                                break

                    new_tokens.append(f_token(expr, 'operand', 'pointer'))

                elif ':OFFSET' in token.tvalue or ':INDEX' in token.tvalue: # example -> A1:OFFSET(
                    depth = 0
                    expr = ''

                    expr += token.tvalue

                    for t in tokens[(index + 1):]:
                        if t.tsubtype == 'start':
                            depth += 1
                        elif t.tsubtype == 'stop':
                            depth -= 1

                        expr += t.tvalue

                        tokens.remove(t)

                        if depth == 0:
                            new_tokens.pop()
                            break

                    new_tokens.append(f_token(expr, 'operand', 'pointer'))


    tokens = new_tokens if new_tokens else tokens

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
        if (stack[-1].tsubtype == "start" or stack[-1].tsubtype == "stop"):
            raise Exception("Mismatched or misplaced parentheses")

        output.append(create_node(stack.pop(), ref))

    # convert to list
    return [x for x in output]

def build_ast(expression, debug = False):
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

                G.add_node(arg1,pos = 1)
                G.add_node(arg2,pos = 2)
                G.add_edge(arg1, n)
                G.add_edge(arg2, n)
            else:
                arg1 = stack.pop()
                G.add_node(arg1,pos = 1)
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
                G.add_node(a,pos = i)
                G.add_edge(a,n)
        else:
            G.add_node(n,pos=0)

        stack.append(n)

    return G,stack.pop()

# Whats the difference between subgraph() and make_subgraph() ?
def subgraph(G, seed):
    subgraph = networkx.DiGraph()
    todo = [(seed,n) for n in G.predecessors(seed)]
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
        todo = [(seed,n) for n in G.predecessors(seed)]
    else:
        todo = [(seed,n) for n in G.successors(seed)]
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


def cell2code(cell, named_ranges):
    """Generate python code for the given cell"""
    if cell.formula:

        debug = False
        # if 'OFFSET' in cell.formula or 'INDEX' in cell.formula:
        #     debug = True
        # if debug:
        #     print 'FORMULA', cell.formula

        ref = parse_cell_address(cell.address()) if not cell.is_named_range else None
        sheet = cell.sheet

        e = shunting_yard(cell.formula, named_ranges, ref=ref, tokenize_range = False)

        ast,root = build_ast(e, debug = debug)
        code = root.emit(ast, context=sheet)

        # print 'CODE', code, ref

    else:
        ast = None
        if isinstance(cell.value, unicode):
            code = u'u"' + cell.value.replace(u'"', u'\\"') + u'"'
        elif isinstance(cell.value, str):
            raise RuntimeError("Got unexpected non-unicode str")
        else:
            code = str(cell.value)
    return code,ast


def prepare_pointer(code, names, ref_cell = None):
    # if ref_cell is None, it means that the pointer is a named_range

    try:
        start, end = code.split('):')
        start += ')'
    except:
        try:
            start, end = code.split(':INDEX')
            end = 'INDEX' + end
        except:
            start, end = code.split(':OFFSET')
            end = 'OFFSET' + end

    def build_code(formula):
        ref = None
        sheet = None

        if ref_cell is not None:
            sheet = ref_cell.sheet

            if not ref_cell.is_named_range:
                ref = parse_cell_address(ref_cell.address())

        e = shunting_yard(formula, names, ref = ref, tokenize_range = False)
        debug = False
        ast,root = build_ast(e, debug = debug)
        code = root.emit(ast, context = sheet, pointer = True)

        return code

    [start_code, end_code] = list(map(build_code, [start, end]))

    # string replacements so that cellmap keys and pointer Range names are coherent
    if ref_cell:
        start_code = start_code.replace("'", '"')
        end_code = end_code.replace("'", '"')

        ref_cell.python_expression = ref_cell.python_expression.replace(code, "%s:%s" % (start_code, end_code))

    return {
        "start": start_code,
        "end": end_code
    }


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
    else: # ~ cell_source is a ExcelCompiler
        cellmap = dict([(x.address(),x) for x in seeds])
        cells = cell_source.cells
        # directed graph
        G = networkx.DiGraph()
        # match the info in cellmap
        for c in cellmap.values(): G.add_node(c)

    # cells to analyze: only formulas
    todo = [s for s in seeds if s.formula]
    steps = [i for i,s in enumerate(todo)]
    names = cell_source.named_ranges

    while todo:
        c1 = todo.pop()
        step = steps.pop()
        cursheet = c1.sheet

        ###### 1) looking for cell c1 dependencies ####################
        # print 'C1', c1.address()
        # in case a formula, get all cells that are arguments
        pystr, ast = cell2code(c1, names)
        # set the code & compile it (will flag problems sooner rather than later)
        c1.python_expression = pystr.replace('"', "'") # compilation is done later

        if 'OFFSET' in c1.formula or 'INDEX' in c1.formula:
            if c1.address() not in cell_source.named_ranges: # pointers names already treated in ExcelCompiler
                cell_source.pointers.add(c1.address())

        # get all the cells/ranges this formula refers to
        deps = [x for x in ast.nodes() if isinstance(x,RangeNode)]
        # remove dupes
        deps = uniqueify(deps)

        ###### 2) connect dependencies in cells in graph ####################

        # ### LOG
        # tmp = []
        # for dep in deps:
        #     if dep not in names:
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
            dep_name = dep.tvalue.replace('$','')

            # this is to avoid :A1 or A1: dep due to clean_pointers() returning an ExcelError
            if dep_name.startswith(':') or dep_name.endswith(':'):
                dep_name = dep_name.replace(':', '')

            # if not pointer, we need an absolute address
            if dep.tsubtype != 'pointer' and dep_name not in names and "!" not in dep_name and cursheet != None:
                dep_name = cursheet + "!" + dep_name

            # Named_ranges + ranges already parsed (previous iterations)
            if dep_name in cellmap:
                origins = [cellmap[dep_name]]
                target = cellmap[c1.address()]
            # if the dep_name is a multi-cell range, create a range object
            elif is_range(dep_name) or (dep_name in names and is_range(names[dep_name])):
                if dep_name in names:
                    reference = names[dep_name]
                else:
                    reference = dep_name

                if 'OFFSET' in reference or 'INDEX' in reference:
                    start_end = prepare_pointer(reference, names, ref_cell = c1)
                    rng = cell_source.Range(start_end)

                    if dep_name in names: # dep is a pointer range
                        address = dep_name
                    else:
                        if c1.address() in names: # c1 holds is a pointer range
                            address = c1.address()
                        else: # a pointer range with no name, its address will be its name
                            address = '%s:%s' % (start_end["start"], start_end["end"])
                            cell_source.pointers.add(address)
                else:
                    address = dep_name

                    # get a list of the addresses in this range that are not yet in the graph
                    range_addresses = list(resolve_range(reference, should_flatten=True)[0])
                    cellmap_add_addresses = [addr for addr in range_addresses if addr not in cellmap.keys()]

                    if len(cellmap_add_addresses) > 0:
                        # this means there are cells to be added

                        # get row and col dimensions for the sheet, assuming the whole range is in one sheet
                        sheet_initial = split_address(cellmap_add_addresses[0])[0]
                        max_rows, max_cols = max_dimension(cellmap, sheet_initial)

                        # create empty cells that aren't in the cellmap
                        for addr in cellmap_add_addresses:
                            sheet_new, col_new, row_new = split_address(addr)

                            # if somehow a new sheet comes up in the range, get the new dimensions
                            if sheet_new != sheet_initial:
                                sheet_initial = sheet_new
                                max_rows, max_cols = max_dimension(cellmap, sheet_new)

                            # add the empty cells
                            if int(row_new) <= max_rows and int(col2num(col_new)) <= max_cols:
                                # only add cells within the maximum bounds of the sheet to avoid too many evaluations
                                # for A:A or 1:1 ranges

                                cell_new = Cell(addr, sheet_new, value="", should_eval='False') # create new cell object
                                cellmap[addr] = cell_new # add it to the cellmap
                                G.add_node(cell_new) # add it to the graph
                                cell_source.cells[addr] = cell_new # add it to the cell_source, used in this function

                    rng = cell_source.Range(reference)

                if address in cellmap:
                    virtual_cell = cellmap[address]
                else:
                    virtual_cell = Cell(address, None, value = rng, formula = reference, is_range = True, is_named_range = True )
                    # save the range
                    cellmap[address] = virtual_cell

                # add an edge from the range to the parent
                G.add_node(virtual_cell)
                # Cell(A1:A10) -> c1 or Cell(ExampleName) -> c1
                G.add_edge(virtual_cell, c1)
                # cells in the range should point to the range as their parent
                target = virtual_cell
                origins = []

                if len(list(rng.keys())) != 0: # could be better, but can't check on Exception types here...
                    for child in rng.addresses:
                        if child not in cellmap:
                            origins.append(cells[child])
                        else:
                            origins.append(cellmap[child])
            else:
                # not a range
                if dep_name in names:
                    reference = names[dep_name]
                else:
                    reference = dep_name

                if reference in cells:
                    if dep_name in names:
                        virtual_cell = Cell(dep_name, None, value = cells[reference].value, formula = reference, is_range = False, is_named_range = True )

                        G.add_node(virtual_cell)
                        G.add_edge(cells[reference], virtual_cell)

                        origins = [virtual_cell]
                    else:
                        cell = cells[reference]
                        origins = [cell]

                    cell = origins[0]

                    if cell.formula is not None and ('OFFSET' in cell.formula or 'INDEX' in cell.formula):
                        cell_source.pointers.add(cell.address())
                else:
                    virtual_cell = Cell(dep_name, None, value = None, formula = None, is_range = False, is_named_range = True )
                    origins = [virtual_cell]

                target = c1


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
                        pystr,ast = cell2code(c2, names)
                        c2.python_expression = pystr
                        c2.compile()

                    # save in the cellmap
                    cellmap[c2.address()] = c2
                    # add to the graph
                    G.add_node(c2)

                # add an edge from the cell to the parent (range or cell)
                if(target != []):
                    # print "Adding edge %s --> %s" % (c2.address(), target.address())
                    G.add_edge(c2,target)

        c1.compile() # cell compilation is done here because pointer ranges might update python_expressions


    return (cellmap, G)
