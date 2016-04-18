import glob
import sys
from datetime import datetime


import os.path
import warnings
from io import BytesIO

from koala.xml.functions import fromstring, safe_iterator
from koala.unzip import read_archive
from koala.excel.excel import read_named_ranges, read_cells
from koala.ast.tokenizer import ExcelParser
from koala.ast.graph import ExcelCompiler



if __name__ == '__main__':

    # files = glob.glob("./example/example3.xlsx")
    files = glob.glob("./data/*.xlsx")

    for file in files:
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

        sys.setrecursionlimit(10000)

        # print 'First evaluation', sp.evaluate('Sheet1!L1')
        # sp.set_value('Sheet1!K1', 100)
        # print 'Second evaluation', sp.evaluate('Sheet1!L1')

        print 'First evaluation', sp.evaluate('Cashflow!G187')
        sp.set_value('InputData!G14', 2025)
        startTime = datetime.now()
        print 'Second evaluation', sp.evaluate('Cashflow!G187')
        print "Evaluation done in %s" % (str(datetime.now() - startTime))

