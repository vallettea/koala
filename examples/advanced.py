
from koala.ExcelCompiler import ExcelCompiler
from koala.Spreadsheet import Spreadsheet
from koala.excellib import xsum

inputs = [
    "Sheet1!A1"
]

outputs = [
    "Sheet1!D1",
    "Sheet1!R1:R4"
]


file = "./examples/advanced.xlsx"

print file

### Graph Generation ###
c = ExcelCompiler(file, ignore_sheets = ['Sheet2'], ignore_hidden = True, debug = True)
c.clean_volatile()
sp = c.gen_graph(inputs= inputs, outputs = outputs)

### Graph Pruning ###
sp = sp.prune_graph()

## Graph Serialization ###
print "Serializing to disk..."
sp.dump(file.replace("xlsx", "gzip"))

### Graph Loading ###
print "Reading from disk..."
sp = Spreadsheet.load(file.replace("xlsx", "gzip"))

### Graph Evaluation ###
sp.set_value('Sheet1!A1', 10)
print 'New D1 value: %s' % str(sp.evaluate('Sheet1!D1'))

# Extracting Ranges (from existing Cells)
r_range = sp.Range('Sheet1!R1:R4')
print 'Created Range from R column', r_range.values


# using Excel functions
print 'SUM([1, 2, 3] =', xsum([1, 2, 3])

# fix a Cell
sp.fix_cell('Sheet1!D1')
sp.set_value('Sheet1!A1', 30)
print 'A1 = 30, but D1 was fixed ==> D1 =', sp.evaluate('Sheet1!D1')

# free a Cell
sp.free_cell()
print 'D1 should eval', sp.cellmap['Sheet1!D1'].should_eval
print 'A1 = 30, but D1 was freed ==> D1 =', sp.evaluate('Sheet1!D1')

# change formulas
sp.set_formula('Sheet1!D1', 'Sheet1!A1 * 1000')
print 'D1 formula was changed to A1 * 1000 ==> D1 =', sp.evaluate('Sheet1!D1')



