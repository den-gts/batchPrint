# -*-coding: utf-8 -*-
from win32com.client import Dispatch
from win32com.universal import com_error
import misc
import os, sys, subprocess, logging

OUTPUT_FORMAT = ".tif"
KompasExt = {".cdw": "Document2D",
             ".spw": "SpcDocument"}
logger = logging.getLogger("main.rasterer")

class RastererException(Exception):
    pass


class FileTypeException(RastererException):
    pass


class PathException(RastererException):
    pass


def runHiddenKompas():
        Kompas = Dispatch("KOMPAS.Application.5")
        #Kompas.Visible = False
        return Kompas


def __rasterKompasFile(doc, outputFile, RasterFormatParam):
    logger.info(u"Расторизация в %s" % outputFile)
    success = doc.SaveAsToRasterFormat(outputFile, RasterFormatParam)
    if not success:
        raise RastererException('Error during rasterization of %s' % outputFile)


def rasterKompasFile(kompasObject, inputFile, outputFile):
    """
    Функция конвертирования файла КОМПАС-График в формат TIFF c DPI 300
    :param inputFile:
    :param outputFile:
    :raise:
    """
    extension = os.path.splitext(inputFile)[1]
    docType = KompasExt.get(extension, None)
    if not docType:
        raise FileTypeException("Extension %s not supported" % extension)
    doc = getattr(kompasObject, docType)
    if not os.path.exists(inputFile):
        raise PathException("file '%s' don't exists" % inputFile)
    outputPath = os.path.split(outputFile)[0]
    if not os.path.isdir(outputPath):
        raise PathException("path '%s' don't exists" % outputPath)
    try:
        doc.ksOpenDocument(inputFile, 1)
        pageCount = 1
        RasterFormatParam = doc.RasterFormatParam()
        RasterFormatParam.Init()
        RasterFormatParam.Format = 4  # TIFF
        RasterFormatParam.ExtResolution = 300
        if docType == "Document2D":
            colorType = 2
            pageCount = doc.ksGetDocumentPagesCount()
        else:
            colorType = 0
        if docType == "SpcDocument":
            pageCount = doc.ksGetSpcDocumentPagesCount()
        RasterFormatParam.ColorType = colorType
        RasterFormatParam.ColorBPP = 8  # 256 color
        RasterFormatParam.MultiPageOutput = False
        RasterFormatParam.OnlyThinLine = False
        RasterFormatParam.RangeIndex = 0  # all pages

        if pageCount > 1:
            for page in range(1, pageCount + 1):
                RasterFormatParam.pages = page
                outputPage = os.path.splitext(outputFile)
                outputPage = ("(%i)" % page).join(outputPage)
                __rasterKompasFile(doc, outputPage, RasterFormatParam)
        else:
            __rasterKompasFile(doc, outputFile, RasterFormatParam)
    except com_error, err:
        logger.error(err[1].decode('cp1251'))
        raise com_error(err[1].decode('cp1251'))
    finally:
        doc.ksCloseDocument()


def rasterFolder(path, outputFolder=None):
    if not outputFolder:
        outputFolder = path
    if not misc.checkWritePath(outputFolder):
        raise IOError(u"Папка %s не доступна для записи" % outputFolder)
    logger.info(u"Начало растеризации папки %s" % path)

    for p in path, outputFolder:
        if not os.path.isdir(p):
            raise PathException('%s not found' % p)

    Kompas = Dispatch("KOMPAS.Application.5")
    try:
        for f in misc.getFileList(path, KompasExt):
            outputFileName = os.path.split(f)[1]
            outputFileName = os.path.splitext(outputFileName)[0]  # отсекаем расширение
            try:
                rasterKompasFile(Kompas, f, "%s/%s%s" % (outputFolder, outputFileName, OUTPUT_FORMAT))
            except RastererException as err:
                print err[1]
    finally:
        # Если компас запущен, не закрываем
        if not Kompas.Visible:
            Kompas.Quit()
    return True


def PDFtoTIFF(sourceFile, outputFile):
    command = ur'"C:\Program Files\gs\gs9.07\bin\gswin64c.exe" -dSAFER -dBATCH -dNOPAUSE  -dQUIET -r300 ' \
              ur'-sDEVICE=tiff24nc -sCompression=lzw -sOutputFile="%s" "%s"' % (outputFile, sourceFile)
    out, err = subprocess.Popen(command.encode('cp1251'), stdout=subprocess.PIPE, shell=True).communicate()
    if out.strip():
        print out.decode('cp1251')

if __name__ == '__main__':
    strHndl = logging.StreamHandler()
    logger.addHandler(strHndl)
    logger.setLevel(logging.DEBUG)
    Kompas = Dispatch("KOMPAS.Application.5")
    rasterKompasFile(Kompas, u"D:/temp/D/test.cdw", u"D:/Temp/Test.tif")
    Kompas.Quit()