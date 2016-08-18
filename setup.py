from setuptools import setup
from setuptools import Extension
from setuptools import find_packages
from Cython.Build import cythonize

setup(
    name = "koala2",

    version = "0.0.10",

    author = "Ants, open innovation lab",
    author_email = "contact@ants.builders",

    packages = find_packages(),

    include_package_data = True,

    url = "https://github.com/anthill/koala",

    license = "MIT",
    description = "A python module to extract all the content of an Excel document and enable calculation without Excel",

    long_description = open("README.md").read(),

    install_requires = [
        "lxml"
    ],

    # ext_modules = cythonize(["koala/*.py", "koala/ast/*.py"]),


)