from distutils.core import setup

setup(
    name = "koala",

    version = "0.0.1",

    author = "Ants, open innovation lab",
    author_email = "contact@ants.builders",

    packages = ["koala"],

    include_package_data = False,

    url = "http://pypi.python.org/pypi/koala_v001/",

    license = "MIT",
    description = "A blazing fast python module to extract all the content of an Excel document and enable calculation without Excel",

    long_description = open("README.md").read(),

    install_requires = [
        "lxml",
    ],
)