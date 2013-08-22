# -*-coding: utf-8 -*-
import rasterer, BPimage, misc
import os, logging, sys
from win32com.client import Dispatch

logger = logging.getLogger("main")
fileformater = logging.Formatter(u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s')
strmFormater = logging.Formatter(u"# %(levelname)s [%(asctime)s]  %(message)s'")
strmHeandler = logging.StreamHandler()
strmHeandler.setLevel(logging.DEBUG)
strmHeandler.setFormatter(strmFormater)
logger.addHandler(strmHeandler)
logger.setLevel(logging.DEBUG)

def rastrPrintKompasFile(inputPath, max_format, outputFile=None, Kompas=None):
    closeKompas = False
    if Kompas is None:
        Kompas = rasterer.runHiddenKompas()
        closeKompas = True
    if outputFile is None:
        outputFile = os.path.join(os.environ['temp'], misc.getFileName(outputFile) + rasterer.OUTPUT_FORMAT)

    logger.info(u"Печать файла КОМПАС %s" % inputPath)
    rasterer.rasterKompasFile(Kompas, inputPath, outputFile)  # функция растеризации файла компаса
    result = BPimage.autoPrintImage(outputFile, max_format)  # функция печати файла компаса
    if closeKompas and not Kompas.Visible:  # TODO need refactor
        Kompas.Quit()
    return result


def rasterAndPrint(path, max_format="A1"):
    outputFolder = path
    Kompas = None
    if not misc.checkWritePath(outputFolder):
        outputFolder = os.environ['temp']
    extensions = rasterer.KompasExt.keys()
    extensions.extend((".tif", ".tiff"))
    logger.info(u"Начало расторизации и печати папки %s" % path)
    fileList = misc.getFileList(path, extensions)
    try:
        for filePath in fileList:
            if os.path.splitext(filePath)[1] in rasterer.KompasExt.keys():
                if Kompas is None:  # Если компас не запущен, запускаем
                    Kompas = rasterer.runHiddenKompas()
                outputFilePath = os.path.join(outputFolder,
                                              misc.getFileName(filePath) + rasterer.OUTPUT_FORMAT)
                if rastrPrintKompasFile(filePath,
                                        max_format,
                                        outputFilePath,
                                        Kompas):
                    try:
                        fileList.remove(outputFilePath)
                    except ValueError:
                        pass

            else:
                BPimage.autoPrintImage(filePath, max_format)
    finally:
        if Kompas and not Kompas.Visible:  # TODO need refactor
            Kompas.Quit()


def printFolder(path, exts=(".tif", ".tiff"), max_format="A1", recursive=True):
    fl = misc.getFileList(path, exts, recursive)
    for f in fl:
        BPimage.autoPrintImage(f, max_format)


def testPDFtoTIFF():
    fl = misc.getFileList(ur'\\ARHIV\Projects\СКИД.461214.003 МССС-18 (Саттелит)\СКИД.301319.052 Базовый транспортный модуль\СКИД.304121.006 Ниша\pdf', '.pdf', False)
    for f in fl:
        outputFile = os.path.splitext(f)[0] + ".tif"
        logger.info(u"Конвертирование %s в %s" % (f, outputFile))
        rasterer.PDFtoTIFF(f, outputFile)


rasterAndPrint(ur'D:\Temp\D')
#  rasterer.rasterFolder(ur"\\ARHIV\Projects\СКИД.566132.001 Блок АВР")