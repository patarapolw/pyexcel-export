"""
Defines ROOT as project_name/project_name/. Useful when installing using pip/setup.py.
"""

import os
import inspect

MODULE_ROOT = os.path.abspath(os.path.dirname(inspect.getframeinfo(inspect.currentframe()).filename))
TRUE_ROOT = os.path.dirname(MODULE_ROOT)


def module_path(filename):
    return os.path.join(MODULE_ROOT, filename)


def root_path(filename):
    return os.path.join(TRUE_ROOT, filename)
