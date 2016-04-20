import glob
import sys
from datetime import datetime


import warnings
from io import BytesIO

from koala.xml.functions import fromstring, safe_iterator
from koala.ast.tokenizer import ExcelParser
from koala.ast.graph import ExcelCompiler
from koala.ast.excelutils import Cell


if __name__ == '__main__':

    file = "./example/example.xlsx"

    print file        
    startTime = datetime.now()

    c = ExcelCompiler(file, ignore_sheets = ['IHS'])
    print "%s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))

    sp = c.gen_graph()

    sys.setrecursionlimit(10000)
    print '- Eval INPUT', sp.evaluate('INPUT')
    print '- Eval A1' , sp.evaluate('Sheet1!A1')
    print '- Eval RESULT', sp.evaluate('RESULT')
    print 'set_value INPUT <- 2025'
    sp.set_value('INPUT', 2025)

    print '- Eval INPUT', sp.evaluate('INPUT')
    print '- Eval A1', sp.evaluate('Sheet1!A1')
    print '- Eval RESULT', sp.evaluate('RESULT')



