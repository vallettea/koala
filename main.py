
import glob, re, sys
from datetime import datetime


import os.path
import warnings
from io import BytesIO
from multiprocessing import Pool

from koala.reader import read_archive, read_named_ranges, read_cells
from koala import ExcelCompiler, Spreadsheet



def calculate_graph(file): 
    print file
    file_name = os.path.abspath(file)
    
    startTime = datetime.now()

    # try:
    c = ExcelCompiler(file, ignore_sheets = ['IHS'], parse_offsets = True)
    print "%s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))
    
    startTime = datetime.now()
    sp = c.gen_graph(outputs = ["outNPV_HostGovt"], inputs = ["gen_discountRate"])
    print "Gen graph in %s" % str(datetime.now() - startTime)

    # startTime = datetime.now()
    # print "Serializing to disk...", file
    # sp.dump(file_name.replace("xlsx", "gzip"))
    # print "Serialized in %s" % str(datetime.now() - startTime)

    # startTime = datetime.now()
    # print "Reading from disk...", file
    # sp = Spreadsheet.load(file_name.replace("xlsx", "gzip"))
    # print "Red in %s" % str(datetime.now() - startTime)

    print 'First evaluation', sp.evaluate('outNPV_HostGovt')

    # for addr, cell in sp.cellmap.items():
    #     sp.history[addr] = {'original': str(cell.value)}
    prev = sp.evaluate('gen_discountRate')
    print "prev,", prev
    sp.set_value('gen_discountRate', 0)
    sp.set_value('gen_discountRate', prev)

    startTime = datetime.now()
    print 'Second evaluation', sp.evaluate('outNPV_HostGovt')
    print "___Timing___  Evaluation done in %s" % (str(datetime.now() - startTime))

    # except:
    #     print "Error in file " + file

if __name__ == '__main__':

    files = glob.glob("./data/m*.xlsx")
    # import random
    # random.shuffle(files)
    # pool = Pool(processes = 4)
    # pool.map(calculate_graph, files)
    # map(calculate_graph, files)
    calculate_graph(files[0])
 