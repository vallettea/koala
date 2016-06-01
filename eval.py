import pyximport; pyximport.install()

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
from koala.ast.excellib import *

from koala.ast.Range import RangeCore


if __name__ == '__main__':

    # files = glob.glob("./data/*.xlsx")
    # file = "./example/example.xlsx"
    # file = "../engie/data/input/100021224 - Far East - Indonesia - Abadi Gas (Phase 1) - Gas - New Project.xlsx"
    file = "../engie/data/input/100021237 - Latin America - Brazil - Marlim Sul Module 4 - Oil - New Project.xlsx"

    print file

    startTime = datetime.now()
    c = ExcelCompiler(file, ignore_sheets = ['IHS'])
    c.clean_volatile()
    print "___Timing___ %s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))
    sp = c.gen_graph(outputs=["outNPV_Proj"])
    print "___Timing___ Graph generated in %s" % (str(datetime.now() - startTime))
    
    sp = sp.prune_graph(["IA_PriceExportGas"])

    print "Serializing to disk...", file
    sp.dump2(file.replace("xlsx", "gzip").replace("input", "graphs"))

    startTime = datetime.now()
    print "Reading from disk...", file
    sp = Spreadsheet.load2(file.replace("xlsx", "gzip").replace("input", "graphs"))
    print "___Timing___ Graph read in %s" % (str(datetime.now() - startTime))

    # import cProfile
    # cProfile.run('Spreadsheet.load2(file.replace("xlsx", "txt"))', 'stats')

    sys.setrecursionlimit(10000)

    print 'First evaluation', sp.evaluate('outNPV_Proj')

    tmp = sp.evaluate('IA_PriceExportGas')

    startTime = datetime.now()
    sp = sp.prune_graph(["IA_PriceExportGas"])
    print "___Timing___  Pruning done in %s" % (str(datetime.now() - startTime))

    for addr, cell in sp.cellmap.items():
        sp.history[addr] = {'original': str(cell.value)}

    startTime = datetime.now()
    sp.set_value('IA_PriceExportGas', 0)
    print "___Timing___  Reset done in %s" % (str(datetime.now() - startTime))
    sp.set_value('IA_PriceExportGas', tmp)
    startTime = datetime.now()

#     import cProfile
#     cProfile.run("sp.evaluate('outNPV_Proj')", 'stats')

    print 'Second evaluation %s' % str(sp.evaluate('outNPV_Proj'))
    print "___Timing___  Evaluation done in %s" % (str(datetime.now() - startTime))

    # print 'CONV DOLLAR 2', sp.eval_ref('conv_Dollar_Real')
    # print 'AC209', sp.eval_ref("Calculations!AC209")
    # print 'AC219', sp.eval_ref("Calculations!AC219")

    # print 'Test', RangeCore.apply('divide',RangeCore.apply('multiply',sp.eval_ref("Calculations!AC209"),sp.eval_ref('conv_Dollar_Real'),(219, 'AC')),4,(219, 'AC'))

    counting = False

    # counting differences
    if counting:
        new_history = {}
        mini = float("inf")

        for addr, item in sp.history.items():
            if 'new' in item and 'original' in item:

                cell = sp.cellmap[addr]

                ori_value = item['original']
                new_value = item['new']

                if is_number(ori_value) and is_number(new_value) and abs(float(ori_value) - float(new_value)) > 0.001:
                    mini = min(sp.history[addr]["priority"], mini)
                    new_history[addr] = sp.history[addr]

        print 'NB different', len(new_history)
        print 'Mini', mini

        with open('history_dif_tot.json', 'w') as outfile:
            json.dump(sp.history, outfile)
        with open('history_dif.json', 'w') as outfile:
            json.dump(new_history, outfile)

    
