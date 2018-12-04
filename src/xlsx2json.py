
import sys
import os
from jkLib import jkCommonLib

def xlsx2json (_sourceXlsxFileName):
    
    return 1

def main (argv = sys.argv):
    # if len (argv) < 2:
    #     print ('usage : xlsx2json [xlsx file]')
    #     return -1

    # _xlsFileName = GetFullPath(argv[1])
    _xlsFileName = jkCommonLib.GetFullPath("../../sample.xlsx")
    if os.path.isfile(_xlsFileName) == False:
        print ("Error: Cannout find file %s!" % _xlsFileName)
    
    print ('start reading %s file @%s' % (jkCommonLib.GetFileNameFromPath(_xlsFileName), jkCommonLib.GetLocatedPath(_xlsFileName) ))
    
    return 1

if __name__ == "__main__":
    ret = main()
    # sys.exit(ret)

