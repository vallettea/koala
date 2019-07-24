# Koala: use your Excel files efficiently without Excel in python

At [Ants](weareants.fr), we've been working the past 4 months on a exciting and challenging project: how to reproduce the behavior of Microsoft Excel using exclusively Python.
The objective was to be **able to process many Excel files of high complexity, in a relatively reduced time, and without the need of opening Excel**.

## Origins

This project was born from the need of one of our clients, a French gas company. One department is entitled to anticipate the evolution of the global gas market, and identify the appropriate investments.
The process they use is based upon 14k gas fields business plans (in activity or in project), which are able to give the estimated rentability of a field over the next 100 years or so.

Each business plan (i.e. each field) is represented as an Excel table holding about 50k+ cells and 250k+ relations between these cells.
Processing a field requires to operate a dichotomy in order to find the equilibrium of the field, which can lead to dozens of iterations on the same table.

Initially, the whole operation on all the 14k+ files as conducted by our client used to take about a week of computation. Way too much to be efficient.
Our job was to find a solution to speed the process up as far as possible, using Python, as they required.

The latency in this case obviously comes from an excessive use of the Excel software: for each field, an Excel file is opened, modified, operated, and closed. It is then a naturally decision to get rid of Excel once and for all, and make all the operations in Python.
Basically, this means recoding the Excel calculation engine.

## Inspirations

To be fair, we didn't recode Excel entirely. A lot of what we needed was already open-sourced.

### [OpenPyXL](http://openpyxl.readthedocs.io/en/default/)
This is a full API able to read/write/manipulate `.xlsx` and `.xlsm` files.

What's good:
- totally independent of Excel
- gets all the informations (formulas, styles, macros, charts...)

What's bad:
- does not calculate
- gets too much information. We only need cell values and formulas.

### [Pycel](https://github.com/dgorissen/pycel)
This is a very nice and simple tool that compiles basic excel workbooks into graphs. You can find more information in [this blogpost](https://dirkgorissen.com/2011/10/19/pycel-compiling-excel-spreadsheets-to-python-and-making-pretty-pictures/).

What's good:
- the essential spreadsheet structure compiled into a graph
- calculations
- serialization possible

What's bad:
- needs an Excel instance running at least once per file
- only basic formulas can be calculated
- the graph is not easily modified once created


## What is Koala

Mainly, the combination of OpenPyXL and Pycel formed a solid basis of what we needed:
- get all the essential information of the Workbook reading the `.xslx` file with OpenPyXL
- convert that information into a graph, serialize and do the calculation with Pycel

Of course, we had to do a little bit of work to be able to make these libraries work together, but mostly because the Excel files we needed to handle were too complex for what offered Pycel.
So we started to extend the Pycel code, but at one point, we had modified the code so much we decided to create a different project.
That's what we call **Koala**.

In the end, we have a program that is able to:
- read the core information of an .xlsx file
- build a graph that represents the calculation structure of the file
- modify the graph to add/change cells or formulas
- reduce the graph size to speed up the calculations, when possible
- make calculations
- draw a graphic representation of your graph
- save the graph into a .gzip file
- load a graph from a .gzip file
And all that, **without having Excel installed**.

These features allow you to process many Excel files, without opening Excel once, which is great deal of time saved.
Plus, since you're just using Python, you can parallelize as much as you want, and save even more time !

BENCHMARK


## The future

We feel Koala has a great potential. We don't see it as a way to replace Excel, but as a complementary tool to help process large amounts of files.

We would like to see this project help the community.
Although we coded Koala to be as generic as possible, it still has been designed for the needs of our client.
Given the number of Excel files we had to handle (about 14k), and their complexity in terms of structure, we cover most of the usual cases.

But not all Excel functionalities are handled yet.
Some Excel functions are not available, and some bugs might appear in very tricky situations. Some work is needed also in making the library more understandable and usable.

This is why we need some feedback.

If you find situations that are not handled correctly by Koala, please contact us so that we can fix it.
If you're interested in the technical details of this project, you can [contribute](https://github.com/vallettea/koala).
If you are in situation that can be solved by Koala, but you're not sure how to use it, we can help you deal with it.
If you have comments of any kind, feel free to say hi !
