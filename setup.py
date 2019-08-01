"""Setup script."""

from setuptools import setup, find_packages
from os import path
this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()


if __name__ == '__main__':
    setup(
        name="koala2",

        version="0.0.37",

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
            'networkx==2.1',
            'openpyxl==2.5.3',
            'numpy>=1.15.0',
            'Cython==0.28.2',
            'lxml==4.1.1',
            'six==1.11.0',
            'scipy>=1.0.0',
            'python-dateutil==2.8.0',
            'backports.functools_lru_cache==1.5'
        ]
    )
