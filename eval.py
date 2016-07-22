import pyximport; pyximport.install()

import resource

import glob
import sys
from datetime import datetime
import json

import warnings
from io import BytesIO

from koala.tokenizer import ExcelParser
from koala.ExcelCompiler import ExcelCompiler
from koala.Spreadsheet import Spreadsheet
from koala.Cell import Cell
from koala.Range import RangeCore
from koala.ast import *
from koala.excellib import *



sys.setrecursionlimit(30000)
limit = 67104768 # maximum stack limit on my machine => use 'ulimit -Ha' on a shell terminal
resource.setrlimit(resource.RLIMIT_STACK, (limit, limit))

personalized_names = {
    "Dev_Fwd": "Cashflow!H81",
    "Pnt_Fwd": "Cashflow!I81",
    "Entitled_revenue_liquids": "Cashflow!L31:DG31",
    "Entitled_revenue_gas": "Cashflow!L32:DG32",
    "Econ_truncated_operating_trigger": "Cashflow!L56:DG56",
    "Nominal_tax_impact_decom_adj": "Cashflow!L60:DG60",
    "Nominal_to_real_multiplier": "Cashflow!L64:DG64"
}

# inputs = [
#     "gen_discountRate", 
#     "IA_PriceExportOil", 
#     "IA_PriceExportGas",
#     "IA_PriceExportCond"
# ]

# outputs = [
#     "CA_Years", 
#     "outNPV_Proj", 
#     "Dev_Fwd", # Cashflow!H81  
#     "Pnt_Fwd", # Cashflow!I81
#     "year_FID"
# ]

with open('../engie/tests/features.json') as data_file:    
    features = json.load(data_file)

inputs = features["primary_input"] + features["secondary_input"]
outputs = features["primary_output"] + features["secondary_output"]


if __name__ == '__main__':

    input_folder = 'inputs'
    # input_folder = 'other/Norway_Output'
    graph_folder = 'test_graphs'
    # graph_folder = 'temp_graphs'

    file_number = '100021224'

    file = glob.glob("../engie/data/%s/%s*.xlsx" % (input_folder, file_number))[0]

    print file

    # ### Graph Generation ###
    # startTime = datetime.now()
    # c = ExcelCompiler(file, ignore_sheets = ['IHS'], ignore_hidden = True, debug = True)
    # for name, reference in personalized_names.items():
    #     c.named_ranges[name] = reference
    # c.clean_volatile()
    # print "___Timing___ %s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))
    # sp = c.gen_graph(outputs=outputs, inputs = inputs)
    # print "___Timing___ Graph generated in %s" % (str(datetime.now() - startTime))

    # ### Graph Pruning ###
    # startTime = datetime.now()
    # # sp = sp.prune_graph()
    # print "___Timing___  Pruning done in %s" % (str(datetime.now() - startTime))

    # ## Graph Serialization ###
    # print "Serializing to disk...", file
    # sp.dump(file.replace("xlsx", "gzip").replace(input_folder, graph_folder))

    ### Graph Loading ###
    startTime = datetime.now()
    print "Reading from disk...", file
    sp = Spreadsheet.load(file.replace("xlsx", "gzip").replace(input_folder, graph_folder))
    print "___Timing___ Graph read in %s" % (str(datetime.now() - startTime))

    ### Graph Evaluation ###
    print 'First evaluation: outNPV_Proj', sp.evaluate('outNPV_Proj')

    tmp = sp.evaluate('input_CapexDecom')

    # tmp = sp.evaluate('IA_PriceExportGas')

    print 'BE200', sp.cellmap['Calculations!BE200'].value

    history = True
    if history:
        sp.activate_history();
        for addr, cell in sp.cellmap.items():
            sp.history[addr] = {'original': str(cell.value)}

    startTime = datetime.now()

    sp.set_value('input_CapexDecom', 0)
    print 'BE200', sp.cellmap['Calculations!BE200'].value

    print "___Timing___  Reset done in %s" % (str(datetime.now() - startTime))
    sp.set_value('input_CapexDecom', tmp)
    startTime = datetime.now()

    # import cProfile
    # cProfile.run("sp.evaluate('outNPV_Proj')", 'stats')

    # THIS EVAL DOES NOT WORK AND I NEED TO KNOW WHY


    print 'TEST', RangeCore.apply('multiply',RangeCore.apply('substract',sp.eval_ref('totalDecom', ref = (200, 'BE')),xsum(sp.eval_ref("Calculations!L200:Calculations!CN200")),(200, 'BE')),sp.eval_ref('Deprec_UOPRates', ref = (200, 'BE')),(200, 'BE'))


    print 'TEST 2', xsum(sp.eval_ref("Calculations!L200:Calculations!CN200"))
    print 'TEST 3', sp.eval_ref('totalDecom', ref = (200, 'BE'))
    print 'TEST 4', sp.eval_ref('Deprec_UOPRates', ref = (200, 'BE'))


    print 'Second evaluation %s' % str(sp.evaluate('outNPV_Proj'))
    print "___Timing___  Evaluation done in %s" % (str(datetime.now() - startTime))

    # saving differences
    if history:
        print 'Nb Different', sp.count
        
        with open('history_dif.json', 'w') as outfile:
            json.dump(sp.history, outfile)
