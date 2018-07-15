import os

from pyexcel_export.dir import TRUE_ROOT


def getfile(filename, path_type='input'):
    return os.path.join(TRUE_ROOT, 'tests', path_type, filename)
