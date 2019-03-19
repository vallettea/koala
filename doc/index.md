# Koala: how we built it

One of the latest projects we conducted at [Ants](weareants.fr) is about Microsoft Excel, more specifically how to reproduce its behavior using Python.

We explained IN THIS POST the origins of the project and what the library does.
In this post, we are going to explain the work we did.

As stated before, Koala is the combination of 2 libraries: OpenPyXL for reading `.xlsx` files and Pycel for graph construction and calculations.

Most of the work we did is inspired by Pycel. We'll give you some insights on how it works.

## Understanding Pycel
Let's take a quick look into the basics of Pycel.

### Cells
Cells hold an address, and a value and Excel formula when needed.

```
def Cell(address, sheet=None, value=None, formula=None):
...
```

### CellRanges
CellRanges only holds the addresses of the Cells it refers to. It is able to build this list from an address such as `A1:C1`.

```
def CellRange(self, address, sheet=None):
...
```

### Building the local formula
Cells can hold also a python version of the Excel formula. This code is generated using the `cell2code()` function. The Excel formula is transfomed  into a local temporary graph using the [RPN notation](https://en.wikipedia.org/wiki/Reverse_Polish_notation). Then the python expression is reconstructed, compiled, and stored in the Cell under `my_cell.python_expression` and `my_cell.compiled_expression`.

### Building the global graph
The graph is built from a seed Cell.
The local graph is constructed, allowing access to the arguments of the Excel formula. All arguments representing an Excel Cell or Range are built and the corresponding nodes and edges are created and added to the graph, in the way where arguments are the parents of the Cell holding the formula.
If an argument Cell holds a formula, it is added to the todo list. This method can be seen as climbing up the graph and building the parent Cells at each node.
The process runs until this todo list is empty. Then your graph is complete.

### Calculations
A calculation in Pycel is done in 2 steps:

1) `set_value(input_cell, new_value)`: this function sets the value of a Cell to a new value. This triggers a chain of `reset()` calls on all the original Cells children, recursively.
The `reset` is important because it sets each reset Cell to "I need an update" state, represented by a `None` value. Then, all Cells dependent of the input Cell will be reset.

2) `evaluate(output_cell)`: starting from the output Cell, the algorithm will climb up the tree and evaluate all Cells that need to. If a Cell doesn't need evaluation, its value is sent back to the previous node, where the calculation is done.
You can think about it as a climb up to get all non resetted values necessary for the calculation (i.e Cells not impacted by the `set_value()`), and a climb down to calculate at each node its new value, until you reach the output Cell. 


## First steps

## Adapt OpenPyXL reader
OpenPyXL is a full API able to read/write/manipulate `.xlsx` and `.xlsm` files.
The problem we had is that it gets too much information: formulas, styles, macros, charts...
We don't need as many information. **We only need cells' essentials: address, value, formula**.
So what we did what write our own parser, in `reader.py`, that did just that, using the philosophy and tools of OpenPyXL.
This is also convenient because this way you can adjust your Cell reading knowing what you are going to need later. For instance, at this stage we don't build the cells the way OpenPyXL, but as we are going to need them later, in the graph construction process.

Some tools from the original OpenPyXL code are situated in the `openpyxl` folder.

```
c = ExcelCompiler(file)
```

### Adapt Pycel structure
The core of Koala is inspired by Pycel, which does most of the work we need.
But, as we are going to see later on in this post, a lot of its functionalities were coded for relatively simple use cases, which is definitely not our case. We had to find a project structure more adapted to the complexity of the project
The following structure is the fruit of our many reflexions while working on the project, and might still evolve in the future.

```
ast/ // holds all AST logic, that handles the graphs
    __init__.py // tools to generate graphs
    astnodes.py // classes of AST nodes
Cell.py // our Cell class
ExcelCompiler.py // class that holds cells read by our reader
ExcelError.py // class representing Excel errors
excellib.py // Excel functions available in Python
Range.py // our Range class
reader.py // reads Excel files and outputs an ExcelCompiler instance
serializer.py // tools to serialize / load graphs
Spreadsheet.py // our Spreadsheet class
```

### Adding missing functions

Adding Excel functions in Pycel is very easy. You just need to take a missing function, look up its signature in [Excel reference](https://support.office.com/en-us/home), write it in Python, and put it in `excellib.py`.
In Koala, it was just as easy, since we used the same philosophy as Pycel.

We added more than 20 functions to be able to calculate our sheets, but also had to modify to some extent the existant ones, so that they match with the Koala code.

For now, there are 33 functions coded in Koala, which covers most of the usual cases. 
But there is still a lot of work to do, since there are more than 400 functions in Excel.

### Reducing graph size

This is an important topic.
By reducing graphs, you obviously reduce its size, but also the evaluation time.

When we think about it, in a big spreadsheet if we know exactly what inputs we want to modify, and what output we want to check, we don't need the cells that aren't involved in the calculation.
So the only cells needed are the cells between the inputs and the outputs.

Pycel already allows us to choose which output cell the graph will be built from.
We added the possibility to also filter the graph according to our inputs. This is called **pruning**.

First, we gather all our inputs children cells into a list, `dependencies`.
Then we build a new graph from scratch. We climb up the original graph from our outputs, and add an edge to the new graph if the current cell is in `dependencies`. Otherwise, it means the cell is not related to one of the inputs so it can be set to a constant cell (and all its parents discarded).


### Handling Names

## The Range Problem

One of the major issues we faced developing Koala was how to handle Ranges.
In Pycel, the Range class simply holds the list of Cell addresses concerned by the Range. This is okay to handle functions like `SUM` or `MIN`, or to outputs the values of the Range.
But this is not enough in a lot of cases, where you need to actually operate Ranges not as simple lists.

```
in Cell C3: A1:A3 + B1:B3
```



### Examples of Excel particularities

Excel basic operations (+, -, *, /, ...) on Ranges are weird.
They are not term by term operations, neither matricial operations.
Basic operations on Ranges output unusual results.

#### Range basic operations
Figures

This is because Excel chooses to operate only the Cells associated to the calling Cell, meaning the Cells that are inlined horizontally or vertically with the Cell holding the formula. If one of the Ranges is not associated to the calling Cell, it outputs a `#VALUE!` error.

This makes sense if you consider Excel as a table, which it basically is, but not so much if you forget about the GUI of Excel, which we do with Koala.

#### Sumproduct
Figures
Some functions such as [`SUMPRODUCT`](https://support.office.com/en-us/article/SUMPRODUCT-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) have still a differente behavior.
Any Range operation inside the `SUMPRODUCT` function is a term by term operation. Then the result is just the sum of the operated Range.

This works independantly of if the Ranges are associated to the calling Cell.


We needed to have a class that would fit these special behaviors, while being adapted to the graph philosophy.


### The Range class

#### Definition

Ranges are `dict`: `(row, col)` => `Cell`.
They are built from a range address, `A1:A3` for instance, or from a list of addresses, `[A1, A2, A3]`.
Basically, a Range holds references to all the Cells objects specified by the address or the list.

The values of the Range are the list of its Cell values, and are accessible with `my_range.values`.

It is important to notice that the Range class can be specific to the spreadsheet, meaning it will have a reference the cellmap, or independent. In this case, we need to specify the values of the cells "manually".

#### Position in the graph
To be able to use Ranges simply, we needed to find a proper place for these instances in the graph.

On graph generation, when a Range is found, 2 objects are created: the Range itself, and a virtual Cell. The virtual Cell has the Range address as address, and holds the Range object as its `__value`.
Hence, the Cell `value` property is as such:

```
@property
    def value(self):
        if self.__is_range:
            return self.__value.values
        else:
            return self.__value
```

Then the virtual Cell representing the Range and all Cells concerned by the Range are linked in the graph, the origin Cells being the parents, and the virtual Cell being the child.


Image

#### Operations

To be able to process Excel-like operations, we still needed to be able to find out whether the Ranges were associated to the calling Cell. This requires to pass the calling Cell as reference.

So we added the `Range.apply()` method.
```
@staticmethod
    def apply(func, first, second, ref = None):

        is_associated = False

        if ref:
            if isinstance(first, RangeCore):
                if first.length == 0:
                    first = 0
                else:
                    is_associated = RangeCore.find_associated_cell(ref, first) is not None
            elif second is not None and isinstance(second, RangeCore):
                if second.length == 0:
                    second = 0
                else:
                    is_associated = RangeCore.find_associated_cell(ref, second) is not None
            else:
                is_associated = False

            if is_associated:
                return RangeCore.apply_one(func, first, second, ref)
            else:
                return RangeCore.apply_all(func, first, second, ref)
        else:
            return RangeCore.apply_all(func, first, second, ref)
```

The `func` argument is the type of operation (`addition`, `division`, ...), `first` and `second` are the actual arguments, and `ref` is the calling Cell address.
`Range.apply()` can be used on Ranges or non Ranges, and chooses to use `Range.apply_one()` (operation on a single cell) or `Range.apply_all()` (operation on all Cells, term by term), depending on if the Ranges are associated to the calling Cell.
The output of `Range.apply_all()` is also a Range, to be able to chain these operations.

## Pointer functions

Pointer functions are Excel functions that output a reference to a cell. The most common examples are `OFFSET` and `INDEX`. These functions can be problematic because they might potentially output any cell in the spreadsheet.
*We are going to use `OFFSET` as the example in this article, but this works pretty much the same with `INDEX` or other pointer functions.*

```
OFFSET(A1, 1, 2) => C2 // we offset A1 by 1 vertically, and 2 horizontally
```

### The `reset()` chain issue

The first problem that comes to mind is link to `reset()`.
As we have seen previously, on `set_value()`, the graph is browsed from the specified cell, and all subsequent children are reset with a `need_update` flag, to indicate the future `evaluate()` that this cell needs evaluation.

But with pointer functions, some cells aren't reset. In the example above, if we call `set_value('C2', 1)`, the cell holding `OFFSET(A1, 1, 2)` as formula will not be reset, since it does not have a direct link with `C2`. Then its evaluation will still output the old value of `C2`.

The solution we implemented is to blindly reset all cells holding pointer functions at every `set_value()` call, which can potentially slow down the calculation.
We'll see later how to be more effective with this solution.

### Cleaning pointer functions

Pointer functions can be a problem at the time of reducing the graph. If we decide to use graph reduction (which we want to do), all cells not in the calculation chain will be left aside. In the previous example, `C2` is not in the graph, although it is needed for the calculation. At the time of evaluation, this situation will trigger an error of type `Cell C2 is not in the graph`.

A solution to this issue is to precalculate the outputs of pointer functions.

The idea is pretty simple:
- identify all cells that have at least one pointer function in their formula
- evaluate each pointer function found
- replace the pointer function in the formula by its evaluation
- generate a reduced graph

```
Formula: OFFSET(A1, 1, 2) + OFFSET(A1, 1, 3)

OFFSET(A1, 1, 2) => C2
OFFSET(A1, 1, 3) => D2

Cleaned formula: C2 + D2
```

Once all pointers have been cleaned and the reduced graph generated, we are assured that all needed cells are present in the graph.


### Cleaning = fixing

Cleaning pointer functions is useful to be able to reduce the graph size without "forgetting" cells. But as you may have noticed, it leads to another issue: what if your pointer function varies ?

Assume the following case:
```
B1 = 2
Formula: OFFSET(A1, 1, B1) => C2
Cleaned formula: C2
```

So now, our cell only holds `C2` as formula, so our cell value is the value of `C2` cell.
But if we actually modify `B1` value to `3`, the expected output should be `OFFSET(A1, 1, 3) => D2`. But in reality, the output is still `C2`. Why ? Because by cleaning the pointers, we actually fixed their outputs.

The cleaning process introduced fixed nodes in the graph, which is a serious issue because our evaluations are potentially compromised.

In this case of pointer functions with varying arguments, we have no other choice than to not clean the graph. Meaning we have to use a complete graph.


### Pointer Ranges

Pointer functions really need special attention. Combined with `Ranges`, they raise another problem which needs to be addressed with care.

Consider the following `name`:

```
my_name => A1:OFFSET(A1, 0, 9)
```

`my_name` represents actually the Range `A1:A10`.
If you decide to clean the pointer functions, well, no problem here. `A1:OFFSET(A1, 0, 9)` becomes `A1:A10`, and is parsed as a `Range`.
But if you can't clean (as we have seen before), you're going to have to parse `A1:OFFSET(A1, 0, 9)` as a `Range`, which you can't with regular Ranges.

Being more generic, how should we parse a `Range` such as `OFFSET(A1, B1, C1):OFFSET(A2, B2, C2)` ?
We need to introduce a new `Range` concept: the Pointer Range.

As we have seen before, a `Range` holds a list of `Cell` references.
The main problem with Pointer Ranges is that you don't know a priori the start and end of the `Range`, due to potential varying arguments.
This means you can't create this list of `Cell` references until you actually know the complete extent of that list, which will be at evaluation. 

So a Pointer Range should be a hollow Range that does not hold anything but the formulas to evaluate the start/end `Cell` references. These formulas will be used on evaluation to create the `Cell` reference list and actually build the `Range`.

But let's see that in practice.

1) Detecting and preparing Pointer Ranges

When building the graph, we identify pointer Ranges simply by checking for keywords in the formula.
```
if 'OFFSET' in formula or 'INDEX' in formula:
```

Then we "prepare" the pointer Range. This means we isolate the start part from the end part of the formula, and build python code for each. The codes are stored in a dictionary.

```
reference = prepare_pointer(formula)
// reference = {'start': start_code, 'end': end_code}
```

The `reference` is then used to build the hollow `Range`. This `reference` dictionary is the equivalent of the simple Excel Range address (`A1:A10`) we use for regular `Ranges`.

```
rng = Range(reference)
```

Then we add the new `Range` to a set, to use it later.

```
pointer_ranges.add(rng.name)
```

Synthesizing:
```
if 'OFFSET' in formula or 'INDEX' in formula:
    reference = prepare_pointer(formula)
    rng = Range(reference)
    pointer_ranges.add(rng.name)
```

This first step required to separate the actual building of the `Range` from its constructor. Indeed, building the list of `Cell` references was done in the `Range` constructor. But now we need to build this list on evaluation for pointer Ranges, and on initialization for regular Ranges.
So the `Range` constructor checks if the `Range` is a pointer by checking the `reference` type, and then decides to `build()` or not.

In the `Range` constructor:
```
if type(reference) == dict:
    is_pointer = True

if not is_pointer:
    self.build()
```

2) Evaluating Pointer Ranges

Now that our Pointer Ranges are pre built, we can use them in evaluation. The only thing needed is to build their start/end `references` before evaluating them.

In `Spreadsheet.eval_ref()`:

```
start = eval(rng.reference['start'])
end = eval(rng.reference['end'])
rng.build('%s:%s' % (start, end))
```

Now we have an up-to-date pointer Range that we can evaluate as usual.

It is important to notice though that this is binded to resetting ALL pointer Ranges on `set_value()`. We need to do that because the link between the input and the pointer Range is not always direct.
For instance, with this example:
```
A2 = 2
A1:OFFSET(A1, 0, INDEX(A1, 2, 1)) => A1:A3
```

The `INDEX` part evaluates to `A2`, which holds `2`. In this case, we have an indirect link between an input, `A2`, and the pointer Range. If we `set_value(A2, 3)`, the pointer Range will need to be updated (it will hold more `Cells`), but its `need_update` flag will never be set to `True` because the link is indirect.

For this reason, it is conservative to blindly reset all pointer Ranges on `set_value()`.


### Detection of alive pointers

We have seen that cleaning pointer functions is mandatory to be able to reduce effectively the graph size, but that it can't be done without risk if some pointer functions have varying arguments.
Say you know exactly what inputs you're going to be modifying. This means all other cells won't be affected manually.

In this situation, we might want to know if there are pointer functions that are affected by our inputs. Let's call these pointers **alive pointers**. If there are none, we can `clean_pointers()` safely, since all our pointers are independent of our inputs.

This is done in 2 steps.

1) Gather all arguments of pointer functions

This is done by climbing up the tree from our outputs, and gathering the `Cells` that have pointer functions in their formula.
We then reconstruct each local formula, and use the Reverse Polish Notation structure to find the arguments and sub arguments of the pointer functions.
These arguments are then stored in a `pointer_arguments` list.

2) Check these arguments aren't affected by our inputs

From all our inputs, we climb down the tree and check at each `Cell` if its address is in `pointer_arguments`. If it's the case, then the pointer function is considered "alive", meaning it is affected by one of the inputs.

```
def detect_alive(self, inputs = None, outputs = None):
    pointer_arguments = self.find_pointer_arguments(outputs)

    if inputs is None:
        inputs = self.inputs

    # go down the tree and list all cells that are pointer arguments
    todo = [self.cellmap[input] for input in inputs]
    done = set()
    alive = set()

    while len(todo) > 0:
        cell = todo.pop()

        if cell not in done:
            if cell.address() in pointer_arguments:
                alive.add(cell.address())

            for child in self.G.successors_iter(cell):
                todo.append(child)
  
            done.add(cell)

    self.pointers_to_reset = alive
    return alive
```

If the output of this function is 0, we can safely clean our pointers, and reduce the graph.

If not, we'll have to use a complete graph. But we can now reduce the evaluation time.
Remember the "reset chain issue" ? On `set_value()`, we had to reset all cells holding pointer functions, which on complex graphs can take a while.
Well now we know exactly which pointer cells we need to reset: those containing alive pointers. Which reduces a lot the evaluation time.

## Performance
