"""Setup script."""

from setuptools import setup, find_packages
from os import path
this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()


if __name__ == '__main__':
    setup(
        name="koala2",

        version="0.0.35",

        author="Ants, open innovation lab",
        author_email="contact@weareants.fr",

        packages=find_packages(),

        include_package_data=True,

        url="https://github.com/anthill/koala",

        license="GNU GPL3",
        description=(
            "A python module to extract all the content of an Excel document "
            "and enable calculation without Excel"
        ),
        long_description_content_type='text/markdown',
        long_description=long_description,

        install_requires=[
            'networkx >= 2.1',
            'openpyxl >= 2.5.3',
            'numpy >= 1.14.2',
            'Cython >= 0.28.2',
            'lxml >= 4.1.1',
            'six >= 1.11.0',
        ]
    )
