"""Setup script."""

from io import open
from setuptools import setup, find_packages
from os import path

this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()


if __name__ == '__main__':
    setup(
        name="koala2",

        version="0.0.36",

        author="Ants, open innovation lab",
        author_email="vallettea@gmail.com",

        packages=find_packages(),

        classifiers=[
            "Programming Language :: Python :: 2.7 :: 3",
            "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
            'Operating System :: Microsoft :: Windows',
            'Operating System :: MacOS :: MacOS X',
        ],

        include_package_data=True,

        url="https://github.com/vallettea/koala",

        license="GNU GPL3",
        description=(
            "A python module to extract all the content of an Excel document and enable calculation without Excel"
        ),
        long_description_content_type='text/markdown',
        long_description=long_description,

        keywords=['xls',
            'excel',
            'spreadsheet',
            'workbook',
            'data analysis',
            'analysis'
            'reading excel',
            'read excel',
            'excel formula',
            'excel formulas',
            'excel equations',
            'excel equation',
            'formula',
            'formulas',
            'equation',
            'equations',
            'timeseries',
            'time series',
            'research',
            'scenario analysis',
            'modelling',
            'model'],


        install_requires=[
            'networkx >= 2.4',
            'openpyxl >= 3.0.3',
            'numpy >= 1.14.2',
            'numpy-financial>=1.0.0',
            'Cython >= 0.29.15',
            'lxml >= 4.5.0',
            'scipy>=1.0.0',
            'python-dateutil>=2.8.0'
        ]
    )
