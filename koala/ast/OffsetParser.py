import re
from collections import OrderedDict

from koala.ast.excelutils import split_address, col2num, index2addres, resolve_range

class OffsetParser(object):
    
    def __init__(self, cells, named_ranges):

        named_ranges_values = {}
        for k,v in named_ranges.items():
            k = k.replace(" ","")
            if "OFFSET" not in v:
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

    def parseOffsetArg(self, arg):
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
        argx = offset.group(2)
        argy = offset.group(3)
        sh, col, row = split_address(offset.group(1))
        colbis = col2num(col) + self.parseOffsetArg(argy)
        rowbis = int(row) + self.parseOffsetArg(argx)
        return index2addres(colbis, rowbis, sh)


    def parseOffsets(self, formula):
        if formula in self.cache.keys():
            return self.cache[formula]
        if re.findall("OFFSET\((.+?),(.+?),(MATCH.+?)\)", formula): 
            offsets = re.subn("OFFSET\((.+?),(.+?),(.+?\))\)", self.shift, formula)
        else:
            offsets = re.subn("OFFSET\((.+?),(.+?),(.+?)\)", self.shift, formula)
        # format with one sheet only
        splits = re.findall("^(.+?)!(.+?):(.+?)!(.+?)$", offsets[0])
        if len(splits) > 0:
            start = splits[0][1]
            end = splits[0][3]
            # check no inversion for rows
            # todo: do it for cols
            start_int = int(re.search("\w+(\d+)", start).group(1))
            end_int = int(re.search("\w+(\d+)", end).group(1))
            if end_int < start_int:
                tmp = end
                end = start
                start = tmp
            result = splits[0][0] + "!" + start + ":" + end
        else:
            result = offsets[0]
        # cache formula
        self.cache[formula] = result
        return result