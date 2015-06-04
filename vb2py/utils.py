import os
import sys

# TODO: use modulePath everywhere.
# This means all data will have to always live under that directory as well.
def rootPath():
    """Return the root path"""
    return os.path.join(os.path.abspath(__file__).split("vb2py")[0], "vb2py")

def modulePath():
    """Return the module path"""
    return os.path.abspath(os.path.join(__file__, ".."))


def relativePath(path):
    """Return the path to a file"""
    return os.path.join(rootPath(), path)
# -- end -- << Utilities >>
