import sys
import os

from . import jkString

def GetFullPath (_path):
    return os.path.normpath(os.path.abspath(_path))

def GetFileNameFromPath (_path):
    return os.path.normpath(os.path.basename(_path))

def GetLocatedPath (_path):
    return os.path.dirname(_path)