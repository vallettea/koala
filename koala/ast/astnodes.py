from __future__ import print_function
# cython: profile=True

from openpyxl.compat import unicode

from koala.excellib import FUNCTION_MAP, IND_FUN
from koala.utils import is_range, split_range, split_address, resolve_range
from koala.ExcelError import *


def to_str(my_string):
    # `unicode` != `str` in Python2. See `from openpyxl.compat import unicode`
    if type(my_string) == str and str != unicode:
        return unicode(my_string, 'utf-8')
    elif type(my_string) == unicode:
        return my_string
    else:
        try:
            return str(my_string)
        except:
            print('Couldnt parse as string', type(my_string))
            return my_string
    # elif isinstance(my_string, (int, float, tuple, Ra):
    #     return str(my_string)
    # else:
    #     return my_string


class ASTNode(object):
    """A generic node in the AST"""

    def __init__(self,token, debug = False):
        super(ASTNode,self).__init__()
        self.token = token
        self.debug = debug
    def __str__(self):
        # if type(self.token.tvalue) == unicode:

        return self.token.tvalue
    def __getattr__(self,name):
        return getattr(self.token,name)

    def children(self,ast):
        args = ast.predecessors(self)
        args = sorted(args,key=lambda x: ast.node[x]['pos'])
        return args

    def parent(self,ast):
        args = list(ast.successors(self))
        return args[0] if args else None


    def find_special_function(self, ast):
        found = False
        current = self

        special_functions = ['sumproduct']
        # special_functions = ['sumproduct', 'match']
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

    def has_ind_func_parent(self, ast):

        if self.parent(ast) is not None and self.parent(ast).tvalue in IND_FUN:
            return True
        else:
            return False


    def emit(self,ast,context=None, pointer = False):
        """Emit code"""
        self.token.tvalue


class OperatorNode(ASTNode):
    def __init__(self, args, ref, debug = False):
        super(OperatorNode,self).__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell
        self.debug = debug
        # convert the operator to python equivalents
        self.opmap = {
                 "=":"==",
                 "&":"+",
                 "":"+" #union
                 }

        self.op_range_translator = {
            "*": "multiply",
            "/": "divide",
            "+": "add",
            "-": "substract",
            "^": "power",
            "==": "is_equal",
            "<>": "is_not_equal",
            ">": "is_strictly_superior",
            "<": "is_strictly_inferior",
            ">=": "is_superior_or_equal",
            "<=": "is_inferior_or_equal"
        }

    def emit(self,ast,context=None, pointer = False):
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
                return "self.eval_ref(%s, ref = %s)" % (','.join([a.emit(ast,context=context) for a in args]), self.ref)


        if self.ttype == "operator-prefix":
            return 'RangeCore.apply_one("minus", %s, None, %s)' % (args[0].emit(ast,context=context), to_str(self.ref))

        if op in ["+", "-", "*", "/", "^", "==", "<>", ">", "<", ">=", "<="]:
            is_special = self.find_special_function(ast)
            call = 'apply' + ('_all' if is_special else '')
            function = self.op_range_translator.get(op)

            arg1 = args[0]
            arg2 = args[1]

            return "RangeCore." + call + "(%s)" % ','.join(['"'+function+'"', to_str(arg1.emit(ast,context=context)), to_str(arg2.emit(ast,context=context)), to_str(self.ref)])

        parent = self.parent(ast)

        #TODO silly hack to work around the fact that None < 0 is True (happens on blank cells)
        if op == "<" or op == "<=":
            aa = args[0].emit(ast,context=context)
            ss = "(" + aa + " if " + aa + " is not None else float('inf'))" + op + args[1].emit(ast,context=context)
        elif op == ">" or op == ">=":
            aa = args[1].emit(ast,context=context)
            ss =  args[0].emit(ast,context=context) + op + '(' + aa + ' if ' + aa + ' is not None else float("inf"))'
        else:
            ss = args[0].emit(ast,context=context) + op + args[1].emit(ast,context=context)


        #avoid needless parentheses
        if parent and not isinstance(parent,FunctionNode):
            ss = "("+ ss + ")"

        return ss


class OperandNode(ASTNode):
    def __init__(self,*args):
        super(OperandNode,self).__init__(*args)
    def emit(self,ast,context=None, pointer = False):
        t = self.tsubtype

        if t == "logical":
            return to_str(self.tvalue.lower() == "true")
        elif t == "text" or t == "error":
            val = self.tvalue.replace('"','\\"').replace("'","\\'")
            return to_str('"' + val + '"')
        else:
            return to_str(self.tvalue)

class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range, e.g., A5, B3:C20 or INPUT """
    def __init__(self,args, ref, debug = False):
        super(RangeNode,self).__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell
        self.debug = debug

    def get_cells(self):
        return resolve_range(self.tvalue)[0]

    def emit(self,ast,context=None, pointer = False):
        if isinstance(self.tvalue, ExcelError):
            if self.debug:
                print('WARNING: Excel Error Code found', self.tvalue)
            return self.tvalue

        is_a_range = False
        is_a_named_range = self.tsubtype == "named_range"

        if is_a_named_range:
            my_str = '"%s"' % self.token.tvalue
        else:
            rng = self.tvalue.replace('$','')
            sheet = context + "!" if context else ""

            is_a_range = is_range(rng)

            if self.tsubtype == 'pointer':
                my_str = '"' + rng + '"'
            else:
                if is_a_range:
                    sh,start,end = split_range(rng)
                else:
                    try:
                        sh,col,row = split_address(rng)
                    except:
                        if self.debug:
                            print('WARNING: Unknown address: %s is not a cell/range reference, nor a named range' % to_str(rng))
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
            (parent.tvalue == 'OFFSET' and parent.children(ast)[0] == self) or
            (parent.tvalue == 'CHOOSE' and parent.children(ast)[0] != self and self.tsubtype == "named_range")) or
            pointer):

            to_eval = False

        # if parent is None and is_a_named_range: # When a named range is referenced in a cell without any prior operation
        #     return 'self.eval_ref(%s, ref = %s)' % (my_str, to_str(self.ref))

        if to_eval == False:
            output = my_str

        # OFFSET HANDLER
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s, ref = %s)' % (my_str, to_str(self.ref))
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s, ref = %s)' % (my_str, to_str(self.ref))

        # INDEX HANDLER
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[0] == self):

            # we don't use eval_ref here to avoid empty cells (which are not included in Ranges)
            if is_a_named_range:
                output = 'resolve_range(self.named_ranges[%s])' % my_str
            else:
                output = 'resolve_range(%s)' % my_str

        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s, ref = %s)' % (my_str, to_str(self.ref))
        elif (parent is not None and parent.tvalue == 'INDEX' and len(parent.children(ast)) == 3 and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            output = 'self.eval_ref(%s, ref = %s)' % (my_str, to_str(self.ref))
        # MATCH HANDLER
        elif parent is not None and parent.tvalue == 'MATCH' \
             and (parent.children(ast)[0] == self or len(parent.children(ast)) == 3 and parent.children(ast)[2] == self):
            output = 'self.eval_ref(%s, ref = %s)' % (my_str, to_str(self.ref))
        elif self.find_special_function(ast) or self.has_ind_func_parent(ast):
            output = 'self.eval_ref(%s)' % my_str
        else:
            output = 'self.eval_ref(%s, ref = %s)' % (my_str, to_str(self.ref))

        return output


class FunctionNode(ASTNode):
    """AST node representing a function call"""
    def __init__(self,args, ref, debug = False):
        super(FunctionNode,self).__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell
        self.debug = False
        # map  excel functions onto their python equivalents
        self.funmap = FUNCTION_MAP

    def emit(self,ast,context=None, pointer = False):
        fun = self.tvalue.lower()

        # Get the arguments
        args = self.children(ast)

        if fun == "atan2":
            # swap arguments
            return "atan2(%s,%s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context))
        elif fun == "pi":
            # constant, no parens
            return "pi"
        elif fun == "if":
            # inline the if

            # check if the 'if' is concerning a Range
            is_range = False
            range = None
            childs = args[0].children(ast)

            for child in childs:
                if ':' in child.tvalue and child.tvalue != ':':
                    is_range = True
                    range = child.tvalue
                    break

            if is_range: # hack to filter Ranges when necessary,for instance situations like {=IF(A1:A3 > 0; A1:A3; 0)}
                return 'RangeCore.filter(self.eval_ref("%s"), %s)' % (range, args[0].emit(ast,context=context))
            if len(args) == 2:
                return "%s if %s else 0" %(args[1].emit(ast,context=context),args[0].emit(ast,context=context))
            elif len(args) == 3:
                return "(%s if %s else %s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context),args[2].emit(ast,context=context))
            else:
                raise Exception("if with %s arguments not supported" % len(args))

        elif fun == "array":
            my_str = '['
            if len(args) == 1:
                # only one row
                my_str += args[0].emit(ast,context=context)
            else:
                # multiple rows
                my_str += ",".join(['[' + n.emit(ast,context=context) + ']' for n in args])

            my_str += ']'

            return my_str
        elif fun == "arrayrow":
            #simply create a list
            return ",".join([n.emit(ast,context=context) for n in args])

        elif fun == "and":
            return "all([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        elif fun == "or":
            return "any([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        elif fun == "index":
            if pointer or self.parent(ast) is not None and self.parent(ast).tvalue == ':':
                return 'index(' + ",".join([n.emit(ast,context=context) for n in args]) + ")"
            else:
                return 'self.eval_ref(index(%s), ref = %s)' % (",".join([n.emit(ast,context=context) for n in args]), self.ref)
        elif fun == "offset":
            if pointer or self.parent(ast) is None or self.parent(ast).tvalue == ':':
                return 'offset(' + ",".join([n.emit(ast,context=context) for n in args]) + ")"
            else:
                return 'self.eval_ref(offset(%s), ref = %s)' % (",".join([n.emit(ast,context=context) for n in args]), self.ref)
        else:
            # map to the correct name
            f = self.funmap.get(fun,fun)
            return f + "(" + ",".join([n.emit(ast,context=context) for n in args]) + ")"
