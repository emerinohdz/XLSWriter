import os
from setuptools import setup

def read(file):
    return open(os.path.join(os.path.dirname(__file__), file)).read()

setup(
    name = "XLS Writer",
    version = "0.1",
    author = "Edgar Merino",
    author_email = "emerino@gmail.com",
    description = "Easy handling of XLS files using python",
    license = "Artistic License 2.0",
    packages = ["devpower"],
    long_description = read("README")
)
