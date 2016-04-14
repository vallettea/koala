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
    archive = read_archive(file_name)        

    named_range = read_named_ranges(archive)
    print "%s named ranged parsed in %s" % (str(len(named_range)), str(datetime.now() - startTime))

    startTime = datetime.now()
    cells = read_cells(archive, ignore_sheets = ['IHS'])

    print "%s cells parsed in %s" % (str(len(cells)), str(datetime.now() - startTime))

    c = ExcelCompiler(named_range, cells)
    sp = c.gen_graph()

    print "Serializing to disk...", file
    sp.save_to_file(file_name.replace("xlsx", "pickle"))

if __name__ == '__main__':

    files = glob.glob("./example/example2.xlsx")
    pool = Pool(processes = 4)
    pool.map(calculate_graph, files)
    # files = glob.glob("./data/m*.xlsx")
    # pool = Pool(processes = 4)
    # pool.map(calculate_graph, files)
    # map(calculate_graph, files)
    calculate_graph(files[0])
