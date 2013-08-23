# -*-coding: utf-8 -*-
import argparse, os.path as path



def isFolderOrFile(chkPath):
    chkPath = path.normpath(chkPath)
    if path.isdir(chkPath):
        fType = 'folder'
    elif path.isfile(chkPath):
        fType = 'file'
    else:
        raise IOError('%s: Wrong file name or folder path' % chkPath)
    return fType

argParser = argparse.ArgumentParser()
argParser.add_argument("-r", action='store_true', help='Raster file or folder')
argParser.add_argument('-p', action="store_true", help='Print file or folder')
argParser.add_argument('-d', action='store_true', help='Delete temparary files after printing')
argParser.add_argument('input', help='File name or folder path')
argParser.add_argument('output', nargs='?', help='Output file name or output folder path')
grpFormat = argParser.add_argument_group('formats', 'options for select formats affect')
grpFormat.add_argument('-tif',action='store_true', help='TIFF format')
grpFormat.add_argument('-kompas', action='store_true', help='KOMPAS format')
grpFormat.add_argument('-pdf',action='store_true', help='PDF format')
grpFormat.add_argument('-jpg',action='store_true', help='JPG format')
opt = argParser.parse_args(r'-r -d test'.split())
print opt
if getattr(opt, 'input', False):
    inputType = isFolderOrFile(opt.input)
    print '%s is %s' % (opt.input, inputType)

