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

def GetGNUString (_str):
    if IsEmpty(_str) == True:
        return ""

    _ret = ""
    for _tok in _str.split(' '):
        # print (_tok)
        if len(_ret) == 0:
            _ret = _tok.upper()
            # print('-> %s' % _ret)
        elif _tok.startswith('_') == False:
            _ret = _ret + "_" + _tok.upper()
            # print('-> %s' % _ret)
        else:
            _ret = _ret + _tok.upper()
            # print('-> %s' % _ret)          
    return _ret

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
        _valueInt = int(_value)
        _valueFloat = float(_value)

        return _valueInt == _valueFloat
    except:
        return False

def IsFloat (_value):
    try:
        _valueInt = int(_value)
        _valueFloat = float(_value)
        return _valueInt != _valueFloat
    except:
        return False