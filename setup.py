"""Setup script."""

from setuptools import setup, find_packages


if __name__ == '__main__':
    setup(
        name="koala2",

        version="0.0.31",

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

        long_description=open("README.md").read(),

        install_requires=[
            'networkx >= 2.1',
            'openpyxl >= 2.5.3',
            'numpy >= 1.14.2',
            'Cython >= 0.28.2',
            'lxml >= 4.1.1',
            'six >= 1.11.0',
        ]
    )
