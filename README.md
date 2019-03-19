# Koala

[![Build Status](https://travis-ci.org/anthill/koala.svg?branch=master)](https://travis-ci.org/anthill/koala)

Koala converts any Excel workbook into a python object that enables on the fly calculation without the need of Excel.

Koala parses an Excel workbook and creates a network of all the cells with their dependencies. It is then possible to change any value of a node and recompute all the depending cells.

## Get started

### Installation ###

Koala is available on pypi so you can just:

```
pip install koala2
```

alternatively, you can download it and install the last version from github:

```
git clone https://github.com/anthill/koala.git
cd koala
python setup.py install
```

### Basic ###

**Koala is still in early stages of developement and feel free to leave us issues when you encounter a problem.**

#### Graph generation

The first thing you need is to convert your workbook into a graph.
This operation may take some time depending on the size of your workbook (we've used koalo on workbooks containg more than 45 000 intricated formulas).

```
from koala.ExcelCompiler import ExcelCompiler

c = ExcelCompiler("examples/basic.xlsx")
sp = c.gen_graph()
```

If this step fails, ensure that your Excel file is recent and in standalone mode (open it with Excel and save, it should rewrite the file and the resulting file should be three of four times heavier).

#### Graph Serialization

As the previous convertion can be long on big graphs, it is often useful to dump the graph to a file:

```
sp.dump('file.gzip')
```

which can be relaoded later with:

```
sp = Spreadsheet.load('file.gzip')
```


#### Graph Evaluation

You can read the values of some cells with `evaluate`. It will only evaluate the calculation if a parent cell has been modified with `set_value`.

```
sp.set_value('Sheet1!A1', 10)
sp.evaluate('Sheet1!D1')
```

#### Named cells or range

If your Excel file has names defined, you can use them freely:

```
sp.set_value('myNamedCell', 0)
```

### Advanced

#### Compiler options

You can pass `ignore_sheets` to ignore a list of Sheets, and `ignore_hidden` to ignore all hidden cells:

```
c = ExcelCompiler(file, ignore_sheets = ['Sheet2'], ignore_hidden = True)
```

In case you have very big files, you might want to reduce the size of the output graph. Here are a few methods.

#### Volatiles

Volatiles are functions that might output a reference to Cell rather than a specific value, which impose a reevaluation everytime. Typical examples are INDEX and OFFSET.

After having created the graph, you can use `clean_pointers` to fix the value of the pointers to their initial values, which reduces the graph size and decreases the evaluation times:

```
sp.clean_pointers()
```

**Warning:** this implies that Cells concerned by these functions will be fixed permanently. If you evaluate a cell whose modified parents are separated by a pointer, you may encounter errors. 
WIP: we are working on automatic detection of the required pointers.

#### Outputs

You can specify the outputs you need. In this case, all Cells not concerned in the calculation of these output Cell will be discarded, and your graph size wil be reduced.

```
sp = c.gen_graph(inputs = ['Sheet1!A1'], outputs=['Sheet1!D1', Sheet1!D2])
```

#### Pruning inputs

In this case, all Cells not impacted by inputs Cells will be discarded, and your graph size wil be reduced.

```
sp = sp.prune_graph()
```

#### Fix and free Cells

You might need to fix a Cell, so that its value is not reevaluated.
You can do that with:

```
sp.fix_cell('Sheet1!D1')
```

By default, all Cells on which you use `sp.set_value()` will be fixed.

You can free your fixed cells with:

```
sp.free_cell('Sheet1!D1') # frees a single Cell
sp.free_cell() # frees all fixed Cells
```

When you free a Cell, it is automatically reevaluated.

#### Set formula

If you need to change a Cell's formula, you can use:

```
sp.set_formula('Sheet1!D1', 'Sheet1!A1 * 1000')
```

The `string` you pass as argument needs to be written with Excel syntax.

** You will find more examples and sample excel files in the directory `examples`.**

#### Detect alive
To check if you have "alive pointers", i.e, pointer functions that have one of your inputs as argument, you can use:

```
sp.detect_alive(inputs = [...], outputs = [...])
```

This will also change the `Spreadsheet.pointers_to_reset` list, so that only alive pointers are resetted on `set_value()`.

## Origins
This project is a "double fork" of two awesome projects:
- [Pycel](https://github.com/dgorissen/pycel), a python module that generates AST graph from a workbook
- [OpenPyXL](http://openpyxl.readthedocs.io/en/default/), a full API able to read/write/manipulate Excel 2010 files.

The most work we did was to adapt [Pycel](https://github.com/dgorissen/pycel) algorithm to more complex cases that it is capable of. This ended up in modifying some core parts of the library, especially with the introduction of `Range` objects.

As for [OpenPyXL](http://openpyxl.readthedocs.io/en/default/), we only took tiny bits, mainly concerning the reading part. Most of what we took from it is left unchanged in the `openpyxl` folder, with references to the original scripts on [BitBucket](https://bitbucket.org/openpyxl/openpyxl).

This module has been enriched by [Ants](http://WeAreAnts.fr), but is part of a more global project of [Engie](http://www.engie.com/) company and particularly it Center of Expertise in Modelling and Economics Studies.

## Licence

GPL
