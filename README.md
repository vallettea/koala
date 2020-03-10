# Koala

[![Build Status](https://travis-ci.org/vallettea/koala.svg?branch=master)](https://travis-ci.org/vallettea/koala)

Koala converts any Excel workbook into a python object that enables on the fly calculation without the need of Excel.

Koala parses an Excel workbook and creates a network of all the cells with their dependencies. It is then possible to change any value of a node and recompute all the depending cells.

You can read more on the origins of Koala [here](doc/presentation.md). If you are looking for ways to contribute, you can get started [here](doc/contribute.md).

## Get started

### Installation ###

Koala is available on pypi so you can just:

```
pip install koala2
```

alternatively, you can download it and install the last version from github:

```
git clone https://github.com/vallettea/koala.git
cd koala
python setup.py install
```

### Basic ###

**Koala is still in early stages of developement and feel free to leave us issues when you encounter a problem.**

#### Graph generation

The first thing you need is to convert your workbook into a graph.
This operation may take some time depending on the size of your workbook (we've used koala on workbooks containing more than 100 000 intricated formulas).

```
from koala.ExcelCompiler import ExcelCompiler
from koala.Spreadsheet import Spreadsheet

sp = Spreadsheet("examples/basic.xlsx")
```

If this step fails, ensure that your Excel file is recent and in standalone mode (open it with Excel and save, it should rewrite the file and the resulting file should be three of four times heavier).

#### Graph Serialization

As the previous conversion can be long on big graphs, it is often useful to dump the graph to a file:

```
sp.dump('file.gzip')
```

which can be reloaded later with:

```
sp = Spreadsheet.load('file.gzip')
```


#### Graph Evaluation

You can read the values of some cells with `cell_evaluate`. It will only evaluate the calculation if a parent cell has been modified with `cell_set_value`.

```
sp.cell_set_value('Sheet1!A1', 10)
sp.cell_evaluate('Sheet1!D1')
```

#### Named cells or range

If your Excel file has names defined, you can use them freely:

```
sp.cell_set_value('myNamedCell', 0)
```

### Advanced

#### Compiler options

You can pass `ignore_sheets` to ignore a list of Sheets, and `ignore_hidden` to ignore all hidden cells:

```
sp = Spreadsheet(file, ignore_sheets = ['Sheet2'], ignore_hidden = True)
```

In case you have very big files, you might want to reduce the size of the output graph. Here are a few methods.

#### Volatiles

Volatiles are functions that might output a reference to Cell rather than a specific value, which impose a reevaluation every time. Typical examples are INDEX and OFFSET.

After having created the graph, you can use `clean_pointers` to fix the value of the pointers to their initial values, which reduces the graph size and decreases the evaluation times:

```
sp.clean_pointers()
```

**Warning:** this implies that Cells concerned by these functions will be fixed permanently. If you evaluate a cell whose modified parents are separated by a pointer, you may encounter errors.
WIP: we are working on automatic detection of the required pointers.

#### Outputs

You can specify the outputs you need. In this case, all Cells not concerned in the calculation of these output Cell will be discarded, and your graph size will be reduced.

```
sp = sp.gen_graph(inputs=['Sheet1!A1'], outputs=['Sheet1!D1', Sheet1!D2])
```

#### Pruning inputs

In this case, all Cells not impacted by inputs Cells will be discarded, and your graph size will be reduced.

```
sp = sp.prune_graph()
```

#### Fix and free Cells

You might need to fix a Cell, so that its value is not reevaluated.
You can do that with:

```
sp.cell_fix('Sheet1!D1')
```

By default, all Cells on which you use `sp.cell_set_value()` will be fixed.

You can free your fixed cells with:

```
sp.cell_free('Sheet1!D1') # frees a single Cell
sp.cell_free() # frees all fixed Cells
```

When you free a Cell, it is automatically reevaluated.

#### Set formula

If you need to change a Cell's formula, you can use:

```
sp.cell_set_formula('Sheet1!D1', 'Sheet1!A1 * 1000')
```

The `string` you pass as argument needs to be written with Excel syntax.

** You will find more examples and sample excel files in the directory `examples`.**

#### Detect alive
To check if you have "alive pointers", i.e., pointer functions that have one of your inputs as argument, you can use:

```
sp.detect_alive(inputs = [...], outputs = [...])
```

This will also change the `Spreadsheet.pointers_to_reset` list, so that only alive pointers are resetted on `cell_set_value()`.

#### Create from scratch
The graph can also be created from scratch (not by using a file).

```
sp_scratch = Spreadsheet()

sp_scratch.cell_add('Sheet1!A1', value=1)
sp_scratch.cell_add('Sheet1!A2', value=2)
sp_scratch.cell_add('Sheet1!A3', formula='=SUM(Sheet1!A1, Sheet1!A2)')

sp_scratch.cell_evaluate('Sheet1!A3')
```

## Licence

GPL
