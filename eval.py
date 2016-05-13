import glob
import sys
from datetime import datetime
import json

import warnings
from io import BytesIO

from koala.xml.functions import fromstring, safe_iterator
from koala.ast.tokenizer import ExcelParser
from koala.ast.graph import ExcelCompiler, Spreadsheet
from koala.ast.excelutils import Cell
from koala.ast.astutils import *

import cProfile


if __name__ == '__main__':

    # files = glob.glob("./data/*.xlsx")
    # file = "./example/example.xlsx"
    file = "./data/100021224 - Far East - Indonesia - Abadi Gas (Phase 1) - Gas - New Project.xlsx"

    print file        
    startTime = datetime.now()

    c = ExcelCompiler(file, ignore_sheets = ['IHS'], parse_offsets = True)
    c.clean_volatile()
    print "___Timing___ %s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))
    sp = c.gen_graph( outputs=['outNPV_Proj'])
    print "___Timing___ Graph generated in %s" % (str(datetime.now() - startTime))
    
    print "Serializing to disk...", file
    sp.dump(file.replace("xlsx", "gzip"))

    startTime = datetime.now()
    print "Reading from disk...", file
    sp = Spreadsheet.load(file.replace("xlsx", "gzip"))
    print "___Timing___ Graph read in %s" % (str(datetime.now() - startTime))

    sys.setrecursionlimit(10000)
    # print '- Eval INPUT', sp.evaluate('INPUT')
    # print '- Eval A1' , sp.evaluate('Sheet1!A1')
    # print '- Eval RESULT', sp.evaluate('RESULT')
    # print 'set_value INPUT <- 2025'
    # sp.set_value('INPUT', 2025)

    # start_node = find_node(sp.G, 'Cashflow!G187')
    # subgraph = subgraph(sp.G, start_node)
    # print 'SUBGRAPH length', subgraph.number_of_nodes()

    print 'First evaluation', sp.evaluate('outNPV_Proj')

    for addr, cell in sp.cellmap.items():
        sp.history[addr] = {'original': str(cell.value)}

    sp.set_value('gen_discountRate', 0)
    print "-------------"
    sp.set_value('gen_discountRate', 0.7)
    startTime = datetime.now()
    print 'Second evaluation %s for %s' % (str(sp.evaluate('outNPV_Proj')),str(-3656.20567668))
    print "___Timing___  Evaluation done in %s" % (str(datetime.now() - startTime))

    print 'NB different', sp.count
    with open('history.json', 'w') as outfile:
        json.dump(sp.history, outfile)

    # startTime = datetime.now()
    # sp.set_value('InputData!G14', 2025)
    # # cProfile.run("sp.set_value('InputData!G14', 2025)")
    # print "___Timing___  Reset done in %s" % (str(datetime.now() - startTime))
    # startTime = datetime.now()
    # # sp.evaluate('Cashflow!G187')
    # # cProfile.run("sp.evaluate('Cashflow!G187')")
    # print 'Second evaluation', sp.evaluate('Calculations!M200')
    # print "___Timing___  Evaluation done in %s" % (str(datetime.now() - startTime))

