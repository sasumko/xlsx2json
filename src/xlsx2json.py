
import sys
import os
from jkLib import jkCommonLib


from openpyxl import Workbook
from openpyxl import load_workbook


def GetLastKeyColumnIndex (_dict):
    _ret = 0
    for _key, _list in _dict.items():
        for _columnId in _list:
            if _columnId > _ret:
                _ret = _columnId
    return _ret

def GetKeyByColumnIndex (_dict, _columnIndex):
    for _key, _list in _dict.items():
        for _columnId in _list:
            if _list.count(_columnIndex) > 0:
                return _key
    return ""

def GetKeyFieldStartRowIndex (_sheet):
    ## fine next row that its first cell filled
    _rowIndex = 3
    for _row in _sheet.iter_rows(min_row=3, max_row=9999):
        if _row[0].value == None:
            _rowIndex = _rowIndex + 1
            continue
        break
    return _rowIndex


def ParseKeyField (_sheet, _keyRowIndex):
    ## read field
    _dictRet = {}
    _columnIndex = 0    #A = 0
    for _field in _sheet[_keyRowIndex]:
        if _field.value == None:
            break
        _strField = str.strip(_field.value)
        
        if jkCommonLib.IsEmpty(_strField) == True or _strField.startswith('_'):
            _columnIndex = _columnIndex + 1
            continue
        
        if _dictRet.get(_strField) == None:
            _dictRet[_strField] = []
        
        _listColumn = _dictRet[_strField]

        if _listColumn == None:
            #Add it!
            _dictRet[_strField] = [ _columnIndex ]
        else:
            _dictRet[_strField].append(_columnIndex)
        _columnIndex = _columnIndex + 1    

    _listField = list(_dictRet.keys())
    _strFields = "KeyField parse done : "

    for _field in _listField:
        _strFields = _strFields + _field + " "

    print (_strFields)

    return _dictRet 


def ParseSheet (_sheet):
    if jkCommonLib.IsEmpty(_sheet['A1'].value) == True or jkCommonLib.IsEmpty(_sheet['A2'].value):
        return
        
    _nameCellValue = str.strip( _sheet['A1'].value ).lower()
    _typeCellValue = str.strip( _sheet['A2'].value ).lower()
    
    # is valid to parse?
    if _nameCellValue != "name":
        return

    if _typeCellValue != "type":
        return

    ##  read def
    _sheetName = _sheet.title
    _jsonName = jkCommonLib.GetCamelString(_sheet['B1'].value)
    _type = _sheet['B2'].value
    print ('start to parse sheet \"%s\"...' % _sheetName)
    print ('\t>JsonName = %s.json' % _jsonName)
    print ('\t>Type = %s' % _type)

    _dataRowIndex = GetKeyFieldStartRowIndex(_sheet)
    _dictFieldByColumn = ParseKeyField(_sheet, _dataRowIndex)
    _lastColumnIndex = GetLastKeyColumnIndex(_dictFieldByColumn)

    ##  read data by field order
    for _row in _sheet.iter_rows(min_row=_dataRowIndex + 1, max_col=_lastColumnIndex + 1):
        if _row[0].value == None:
            #if first column is none == comment row to be skipped
            continue
        
        _columnIndex = 0
        for _value in _row:
            _key = GetKeyByColumnIndex(_dictFieldByColumn, _columnIndex)
            _valueToStore = _value.value
            if jkCommonLib.IsEmpty(_key) == False and _valueToStore != None:
                print ('key:%s\t\tvalue:%s' % (_key, _valueToStore))

            _columnIndex = _columnIndex + 1
            
    


def xlsx2json (_sourceFile):
    _workbook = load_workbook(_sourceFile)

    print (_workbook.sheetnames)

    for _sheet in _workbook:
        ParseSheet(_sheet)
        
    return 1

def main (argv = sys.argv):
    # if len (argv) < 2:
    #     print ('usage : xlsx2json [xlsx file]')
    #     return -1

    # _xlsFileName = GetFullPath(argv[1])
    _xlsFileName = jkCommonLib.GetFullPath("../../sample.xlsx")
    if os.path.isfile(_xlsFileName) == False:
        print ("Error: Cannout find file %s!" % _xlsFileName)
    
    print ('start reading %s file @%s' 
            % (jkCommonLib.GetFileNameFromPath(_xlsFileName), 
                jkCommonLib.GetLocatedPath(_xlsFileName) )
        )
    
    
    xlsx2json(_xlsFileName)
    

    return 1

if __name__ == "__main__":
    ret = main()
    # sys.exit(ret)

