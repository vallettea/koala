"""Setup script."""

from setuptools import setup, find_packages


if __name__ == '__main__':
    setup(
        name = "koala2",

        version = "0.0.18",

        author = "Ants, open innovation lab",
        author_email = "contact@weareants.fr",

        packages = find_packages(),

        include_package_data = True,

        url = "https://github.com/anthill/koala",

        license = "GNU GPL3",
        description = (
            "A python module to extract all the content of an Excel document "
            "and enable calculation without Excel"
        ),

        long_description = open("README.md").read(),

        install_requires = open('requirements.txt').readlines(),
    )
