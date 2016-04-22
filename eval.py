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

    # files = glob.glob("./data/*.xlsx")
    # file = "./example/example.xlsx"
    file = "./data/100021224 - Far East - Indonesia - Abadi Gas (Phase 1) - Gas - New Project.xlsx"

    print file        
    startTime = datetime.now()

    c = ExcelCompiler(file, ignore_sheets = ['IHS'])
    print "%s cells and %s named_ranges parsed in %s" % (str(len(c.cells)-len(c.named_ranges)), str(len(c.named_ranges)), str(datetime.now() - startTime))

    sp = c.gen_graph()

    sys.setrecursionlimit(10000)
    # print '- Eval INPUT', sp.evaluate('INPUT')
    # print '- Eval A1' , sp.evaluate('Sheet1!A1')
    # print '- Eval RESULT', sp.evaluate('RESULT')
    # print 'set_value INPUT <- 2025'
    # sp.set_value('INPUT', 2025)


    # print 'First evaluation', sp.evaluate('Sheet2!E2')
    # sp.set_value('Sheet2!A1', 10)
    # print 'Second evaluation', sp.evaluate('Sheet2!E2')

    # print '- Eval INPUT', sp.evaluate('INPUT')
    # print '- Eval A1', sp.evaluate('Sheet1!A1')
    # print '- Eval RESULT', sp.evaluate('RESULT')



    print 'First evaluation', sp.evaluate('Cashflow!G187')
    sp.set_value('InputData!G14', 2025)
    startTime = datetime.now()
    print 'Second evaluation', sp.evaluate('Cashflow!G187')
    print "Evaluation done in %s" % (str(datetime.now() - startTime))

