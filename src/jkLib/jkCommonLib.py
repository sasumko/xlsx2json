import sys
import os


def GetFullPath (_path):
    return os.path.normpath(os.path.abspath(_path))

def GetFileNameFromPath (_path):
    return os.path.normpath(os.path.basename(_path))

def GetLocatedPath (_path):
    return os.path.dirname(_path)

def IsEmpty (_str):
    return not bool(_str and _str.strip())

def GetCamelString (_str):
    if IsEmpty(_str) == True:
        return ""

    _ret = ""
    for _tok in _str.split(' '):
        _add = _tok.title()

        _ret = _ret + _add
    return _ret


def IsInt (_value):
    try:
        int(_value)
        return True
    except:
        return False

def IsFloat (_value):
    try:
        float(_value)
        return True
    except:
        return False