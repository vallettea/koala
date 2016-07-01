# Koala

Koala is a python module to replace excel formulas. Koala parses an Excel workbook and creates an AST network of all the needed formula. It is then possible to change any value of a cell and recompute the depending cells.

## Origins
This project is a "double fork" of two awesome projects:
- [Pycel](https://github.com/dgorissen/pycel), a python module that generates AST graph from a workbook
- [OpenPyXL](http://openpyxl.readthedocs.io/en/default/), a full API able to read/write/manipulate Excel 2010 files.

The most work we did was to adapt [Pycel](https://github.com/dgorissen/pycel) algorithm to more complex cases that it is capable of. This ended up in modifying some core parts of the library, especially with the introduction of `Range` objects.

As for [OpenPyXL](http://openpyxl.readthedocs.io/en/default/), we only took tiny bits, mainly concerning the reading part. Most of what we took from it is left unchanged in the `openpyxl` folder, with references to the original scripts on [BitBucket](https://bitbucket.org/openpyxl/openpyxl).

This module has been enriched by [Ants](http://WeAreAnts.fr), but is part of a more global project of [Engie](http://www.engie.com/) company and particularly it Center of Expertise in Modelling and Economics Studies.

## Get started

### Basic ###

#### Graph generation

You can generate your Excel graph using:

```
from koala.ExcelCompiler import ExcelCompiler

c = ExcelCompiler(file)
sp = c.gen_graph()
```

#### Graph Serialization
You can dump the graph of your Excel with
```
sp.dump('file.gzip')
```

Then, you can load your graph with
```
sp = Spreadsheet.load('file.gzip')
```

Once the graph created and loaded, you don't need Excel anymore.

#### Graph Evaluation
```
sp.set_value('Sheet1!A1', 10)
sp.evaluate('Sheet1!D1')
```

#### Names
If your Excel file has names defined, you can use them freely.
```
sp.set_value('myNameCell', 0)
```

### Advanced

#### Compiler options
You can pass `ignore_sheets` to ignore a list of Sheets, and `ignore_hidden` to ignore all hidden cells.
```
c = ExcelCompiler(file, ignore_sheets = ['Sheet2'], ignore_hidden = True)
```

In case you have very big files, you might want to reduce the size of the output graph. Here are a few methods.

#### Volatiles
Volatiles are functions that might output a reference to Cell rather than a specific value, which impose a reevaluation everytime.

You can do that by cleaning volatiles, that is, pre evaluate your volatile functions before actually creating the graph.

**Warning:** this implies that Cells concerned by these functions will be fixed permanently.

#### Outputs
You can select the outputs you need. In this case, all Cells not concerned in the calculation of these output Cell will be discarded, and your graph size wil be reduced.
```
sp = c.gen_graph(outputs=['Sheet1!D1', Sheet1!D2])
```

#### Pruning inputs
You can select the inputs you want to modify. In this case, all Cells not impacted by these inputs Cells will be discarded, and your graph size wil be reduced.
```
sp = sp.prune_graph([Sheet1!A1])
```

#### Fix and free Cells
You might need to fix a Cell, so that its value is not reevaluated.
You can do that with
```
sp.fix_cell('Sheet1!D1')
```

By default, all Cells on which you use `sp.set_value()` will be fixed.

You can free your fixed cells with
```
sp.fix_cell('Sheet1!D1') # frees a single Cell
sp.fix_cell() # frees all fixed Cells
```

When you free a Cell, it is automatically reevaluated.

#### Set formula
If you need to change a Cell's formula, you can use
```
sp.set_formula('Sheet1!D1', 'Sheet1!A1 * 1000')
```

The `string` you pass as argument needs to be written with Excel syntax.

personalized_names

load
dump
set_value
activate_history
evaluate