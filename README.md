## Koala

Koala is a python module to replace excel formulas. Koala parses an Excel workbook and creates an AST network of all the needed formula. It is then possible to change any value of a cell and recompute the depending cells.

This project is a "double fork" of two awesome projects:
- [Pycel](https://github.com/dgorissen/pycel), a python module that generates AST graph from a workbook
- [OpenPyXL](http://openpyxl.readthedocs.io/en/default/), a full API able to read/write/manipulate Excel 2010 files.

The most work we did was to adapt [Pycel](https://github.com/dgorissen/pycel) algorithm to more complex cases that it is capable of. This ended up in modifying some core parts of the library, especially with the introduction of `Range` objects.

As for [OpenPyXL](http://openpyxl.readthedocs.io/en/default/), we only took tiny bits, mainly concerning the reading part. Most of what we took from it is left unchanged in the `openpyxl` folder, with references to the original scripts on [BitBucket](https://bitbucket.org/openpyxl/openpyxl).


This module has been enriched by [Ants](http://WeAreAnts.fr), but is part of a more global project of [Engie](http://www.engie.com/) company and particularly it Center of Expertise in Modelling and Economics Studies.