# Contribute

If you are looking for a way to improve the behaviour of Koala, this is the place to be. The sections below describe various ways to contribute to this great project.

## Test and report bugs

A bug found is a bug solved. If you find any issues with Koala, please report them in the [issue section](https://github.com/vallettea/koala/issues) on Github.

## Add new Excel functions

According to the [latest documentation](https://support.office.com/en-us/article/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188), there are 476 Excel functions (all functions are listed in `doc/functions.xlsx`). As of this moment, not all of the functions are completed yet. We depend on people like you to help us to define them all. If you encounter one of these unmapped functions, the process below will clarify how to add them.

All the Excel functions are mapped in `koala/excellib.py`. If a function doesn't work, it means that it isn't defined in this file. To add a new function, follow the following steps:
1. Add the name in the list `IND_FUN` (order alphabetically) in `koala/excellib.py`.
2. Define a new function in the remainder of the file (also order alphabetically). The function name must be equal to the excel function name but lowercase. The arguments also have to be the same as the argument names in Excel. Exceptions are when the function takes many unnamed arguments (e.g. `SUM`). An example for the function `EOMONTH` can be found [here](https://github.com/vallettea/koala/blob/62296bdc9e5f42dde6ff72dc436339c07b963b30/koala/excellib.py#L527).
3. Write the code in the function body and return the value that the function should return. Useful helper functions can be found and/or added in `koala/utils.py`. Please also pay attention to handling the inputs. Return the right `ExcelError` in case of faulty inputs.
4. Document the behaviour of this function with doctstrings.
4. Add tests to `tests/excel/test_functions.py ` (again in alphabetic order). If multiple arguments are defined, each of them should be tested appropriately. An example for `EOMONTH` can be found [here](https://github.com/vallettea/koala/blob/62296bdc9e5f42dde6ff72dc436339c07b963b30/tests/excel/test_functions.py#L871).
5. If new imports are used, add them to `requirements.txt` and `setup.py`.

A full pull request than can be used as an example can be found [here](https://github.com/vallettea/koala/pull/175/files).

If you have completed all these steps, commit the changes to your fork and open a pull request on Github. Good luck!