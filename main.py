import glob
from datetime import datetime


import os.path
import warnings
from io import BytesIO
from multiprocessing import Pool

from koala.xml.functions import fromstring, safe_iterator
from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from koala.ast.tokenizer import ExcelParser
from koala.ast.graph import ExcelCompiler, Spreadsheet


def calculate_graph(file): 
    print file
    file_name = os.path.abspath(file)
    
    startTime = datetime.now()

    c = ExcelCompiler(file, ignore_sheets = ['IHS'])
    print "%s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))
    sp = c.gen_graph()

    startTime = datetime.now()
    print "Serializing to disk...", file
    sp.dump(file_name.replace("xlsx", "gzip"))
    print "Serialized in %s" % str(datetime.now() - startTime)

    startTime = datetime.now()
    print "Reading from disk...", file
    sp = Spreadsheet.load(file_name.replace("xlsx", "gzip"))
    print "Red in %s" % str(datetime.now() - startTime)

if __name__ == '__main__':

    files = glob.glob("./data/m*.xlsx")
    # pool = Pool(processes = 4)
    # pool.map(calculate_graph, files)
    # map(calculate_graph, files)
    calculate_graph(files[0])
 