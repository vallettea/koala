from koala.ExcelCompiler import ExcelCompiler
from koala.Spreadsheet import Spreadsheet

filename = "./examples/basic.xlsx"

print(filename)

### Graph Generation ###
c = ExcelCompiler(filename)
sp = c.gen_graph()

## Graph Serialization ###
print("Serializing to disk...")
sp.dump(filename.replace("xlsx", "gzip"))

### Graph Loading ###
print("Reading from disk...")
sp = Spreadsheet.load(filename.replace("xlsx", "gzip"))

### Graph Evaluation ###
sp.set_value('Sheet1!A1', 10)
print('New D1 value: %s' % str(sp.evaluate('Sheet1!D1')))

