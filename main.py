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
from koala.ast.graph import ExcelCompiler


def calculate_graph(file): 
    print file
    file_name = os.path.abspath(file)
    
    startTime = datetime.now()

    c = ExcelCompiler(file, ignore_sheets = ['IHS'])
    print "%s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))
    sp = c.gen_graph()

    print "Serializing to disk...", file
    sp.save_to_file(file_name.replace("xlsx", "pickle"))

if __name__ == '__main__':

    files = glob.glob("./data/*.xlsx")
    pool = Pool(processes = 4)
    pool.map(calculate_graph, files)
    map(calculate_graph, files)
 