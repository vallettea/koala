import pyximport; pyximport.install()

import glob, re
from datetime import datetime


import os.path
import warnings
from io import BytesIO
from multiprocessing import Pool
from collections import OrderedDict

from koala.xml.functions import fromstring, safe_iterator
from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from koala.ast.tokenizer import ExcelParser
from koala.ast.graph import ExcelCompiler, Spreadsheet
from koala.ast.excelutils import *
from koala.ast.excellib import match
from koala.ast.Range import Range


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
        self.cache[formula] = offsets
        result = offsets[0]
        # if ':' in result:
        #     tmp = result.split(":")
        #     result = tmp[0]+':'
        #     if '!' in tmp[1]:
        #         result += tmp[1].split('!')[1]
        #     else:
        #         result += tmp[1]
        return result


def calculate_graph(file): 
    print file
    file_name = os.path.abspath(file)
    
    startTime = datetime.now()

    try:
        c = ExcelCompiler(file, ignore_sheets = ['IHS'])
        print "%s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))
        



        parser = OffsetParser(c.cells, c.named_ranges)

        offsets = []
        for k,v in c.named_ranges.items():
            if 'OFFSET' in v:
                offsets.append(v)

        for cell in c.cells.values():
            if cell.formula and 'OFFSET' in cell.formula:
                offsets.append(cell.formula)


        for formula in offsets:
            parser.parseOffsets(formula)

        

        # argset = set()
        # for formula in offsets:
        #     all_groups = re.findall("OFFSET\((.+?),(.+?),(.+?)\)", formula)
        #     for groups in all_groups:
        #         # if groups[0] != '0': argset.add(groups[0])
        #         if groups[1] != '0': argset.add(groups[1])
        #         if groups[2] != '0': argset.add(groups[2])


        # print "==================="
        # print argset
        # clean args
        # sp = c.gen_graph(outputs = list(argset))
    except Exception as e:
        print "error " + str(e)
        pass

    # startTime = datetime.now()
    # sp = c.gen_graph()
    # print "Gen graph in %s" % str(datetime.now() - startTime)

    # startTime = datetime.now()
    # print "Serializing to disk...", file
    # sp.dump(file_name.replace("xlsx", "gzip"))
    # print "Serialized in %s" % str(datetime.now() - startTime)

    # startTime = datetime.now()
    # print "Reading from disk...", file
    # sp = Spreadsheet.load(file_name.replace("xlsx", "gzip"))
    # print "Red in %s" % str(datetime.now() - startTime)

if __name__ == '__main__':

    files = glob.glob("./data/*.xlsx")
    import random
    random.shuffle(files)
    # pool = Pool(processes = 4)
    # pool.map(calculate_graph, files)
    map(calculate_graph, files)
    # calculate_graph(files[0])
 