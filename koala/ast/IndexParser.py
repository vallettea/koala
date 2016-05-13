import re
from collections import OrderedDict

from koala.ast.excelutils import split_address, col2num, index2addres, resolve_range
from koala.ast.graph import shunting_yard

class IndexParser(object):
    
    def __init__(self, cells, named_ranges):

        named_ranges_values = {}
        for k,v in named_ranges.items():
            k = k.replace(" ","")
            if "INDEX" not in v:
                if ":" in v:
                    named_ranges_values[k] = v
                else:
                    if v in cells.keys():
                        named_ranges_values[k] = str(cells[v].value)
                    else:
                        named_ranges_values[k] = '0'

        self.cells = cells
        self.named_ranges_values = OrderedDict(sorted(named_ranges_values.items(), key=lambda t: len(t[0]), reverse=True))
        self.cache = {}

        allindex = []
        for k,v in named_ranges.items():
            if 'INDEX' in v:
                allindex.append(v)
        for k,cell in self.cells.items():
            if cell.formula and 'INDEX' in cell.formula:
                allindex.append(cell.formula)

        print "%s index to parse" % str(len(allindex))

        def print_value_tree(ast,addr,indent):
            cell = self.cellmap[addr]
            print "%s %s = %s" % (" "*indent,addr,cell.value)
            for c in ast.predecessors_iter(cell):
                print_value_tree(ast, c.address(), indent+1)

        for formula in allindex:
            print formula
            e = shunting_yard(formula)
            ast,root = build_ast(e)
            print print_value_tree(ast, root, 1)

    def parseIndexArg(self, arg):
        # replace variables
        for k in self.named_ranges_values:
            if k in arg:
                arg = arg.replace(k, self.named_ranges_values[k])
        if "MATCH" in arg:
            def evalMatch(y):
                adds = resolve_range(y.group(2), True)[0]
                vals = [str(self.cells[a].value) for a in adds]
                if y.group(3) != '0':
                    raise Exception("There is no 0 in Match !")
                if y.group(1) in vals:
                    return str(vals.index(y.group(1))+1)
                else:
                    return '0' # dirty but we assume Na in offset arg is zero
            replacedString = re.subn("MATCH\((.+?),(.+?),(.+?)\)", evalMatch, arg)[0].replace(" ", "")
            return eval(replacedString)
        elif "COUNTA" in arg:
            def evalCounta(y):
                cells_address, nb, toto  = resolve_range(y.group(1))
                # compete fonctional version but takes too much time
                def tata(x):
                    if x in self.cells and self.cells[x].value != None:
                        return 1
                    else:
                        return 0
                return str(sum(map(lambda x: tata(x), cells_address)))

            replacedString = re.subn("COUNTA\((.+?)\)", evalCounta, arg)[0].replace(" ", "")
            return eval(replacedString)
        elif "!" in arg:
            sheet_name, position = arg.split("!")
            return int(self.cells[sheet_name+"!"+position].value)
        else:
            try:
                return eval(arg)
            except:
                raise Exception('method embedded in OFFSET formula not implemented')

    
    def shift(self, offset):
        tab = offset.group(1)
        argx = offset.group(2)
        argy = offset.group(3)
        print "INDEX ============"
        print tab, argx, argy
        if tab in self.named_ranges_values:
            tab = self.named_ranges_values[tab]
        print tab
        print "--------"
        print argy
        for name in self.named_ranges_values:
            if name in argy:
                argy = argy.replace(name, self.named_ranges_values[name])
        print argy
        print "--------"
        mat = resolve_range(tab)
        result = mat[0][eval(argy)-1]
        print result
        # sh, col, row = split_address(tab)
        # colbis = col2num(col) + self.parseIndexArg(argy)
        # rowbis = int(row) + self.parseIndexArg(argx)
        # return index2addres(colbis, rowbis, sh)
        return result


    def parseIndexs(self, formula):
        if formula in self.cache.keys():
            return self.cache[formula]
        offsets = re.subn("INDEX\((.+?),(.+?),(.+?)\)", self.shift, formula)
        # format with one sheet only
        # splits = re.findall("^(.+?)!(.+?):(.+?)!(.+?)$", offsets[0])
        # if len(splits) > 0:
        #     start = splits[0][1]
        #     end = splits[0][3]
        #     # check no inversion for rows
        #     # todo: do it for cols
        #     start_int = int(re.search("\w+(\d+)", start).group(1))
        #     end_int = int(re.search("\w+(\d+)", end).group(1))
        #     if end_int < start_int:
        #         tmp = end
        #         end = start
        #         start = tmp
        #     result = splits[0][0] + "!" + start + ":" + end
        # else:
        #     result = offsets[0]
        # # cache formula
        # self.cache[formula] = result
        # return result

        return formula






