import pyximport; pyximport.install()

from koala.tokenizer import ExcelParser
from koala.ExcelCompiler import ExcelCompiler
from koala.Spreadsheet import Spreadsheet
from koala.ast import *
from koala.excellib import *

file = "./example/basic.xlsx"

print file

### Graph Generation ###
c = ExcelCompiler(file)
sp = c.gen_graph()

## Graph Serialization ###
print "Serializing to disk..."
sp.dump(file.replace("xlsx", "gzip"))

### Graph Loading ###
print "Reading from disk..."
sp = Spreadsheet.load(file.replace("xlsx", "gzip"))

### Graph Evaluation ###
sp.set_value('Sheet1!A1', 10)
print 'New D1 value: %s' % str(sp.evaluate('Sheet1!D1'))

