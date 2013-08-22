# -*-coding: utf-8 -*

import win32print, sys
import win32ui, win32gui
import win32con, pywintypes
import os, logging, misc
from PIL import Image, ImageWin
from math import ceil
#
# Constants for GetDeviceCaps
#
#
win32con.DMPAPER_A1 = 271
win32con.DMPAPER_A1_OVERSIZE = 621
win32con.DMPAPER_A2_OVERSIZE = 620

# HORZRES / VERTRES = printable area
#
HORZRES = 8
VERTRES = 10
#
# LOGPIXELS = dots per inch
#
LOGPIXELSX = 88
LOGPIXELSY = 90
#
# PHYSICALWIDTH/HEIGHT = total area
#
PHYSICALWIDTH = 110
PHYSICALHEIGHT = 111
#
# PHYSICALOFFSETX/Y = left / top margin
#
PHYSICALOFFSETX = 112
PHYSICALOFFSETY = 113

# Xerox 5016 trays
XEROX_A3_TRAY = 1
XEROX_A4_TRAY = 15
PAPER_FORMATS = ("A4", "A3", "A2", "A1")
PRINTERS = {"Xerox 5016":
            {
                "Name": "Xerox WorkCentre 5016 at //dlink-01520C/dlk-01520C-U1",
                "PaperFormats": {"A4": XEROX_A4_TRAY, "A3": XEROX_A3_TRAY}
            },
            "Plotter":
            {
                "Name": "HP Designjet 500 24 by HP",
                "PaperFormats": ("A4", "A3", "A2", "A1")
            }
            }

KOMPAS_EXTENSIONS = (".cdw", ".spw")

logger = logging.getLogger("main.print")


class printImageException(Exception):
    pass


class PaperFormatException(printImageException):
    pass


def getImageSizeInMM(img):
    """
    Функция определяет физический размер требуемый для печати изходя из DPI и размера изображения в пикселях.
    :param img: объект, получаемый из PIL.Image
    :return: list(width,hieght) размер изображения в милиметрах
    """
    dpi = img.info['dpi']
    return map(lambda size, lDPI: int(ceil(float(size) / lDPI * 25.4)), img.size, dpi)


def getImagePaperFormat(img):
    # Paper formats
    """
    Функция определяет формат бумаги, требуемый для печати вполный размер. Исходные данные для определения формата
    бумаги - размер изображения в пикселях и количество точек на дюйм(DPI).
    :param img: объект, получаемый из PIL.Image
    :return: формат бумаги в формате A0,A1,A2,A3,A4. None при неудаче при определении формата
    """
    A = [(841, 1189)]
    OFFSET = 25
    for x in range(1, 5):
        prvHeight = A[x - 1][1]
        prvWidth = A[x - 1][0]
        A.append((prvHeight / 2, prvWidth))

    imgSizeMM = getImageSizeInMM(img)
    if imgSizeMM[0] > imgSizeMM[1]:
        imgSizeMM.reverse()

    for pStandard in A:
            delta = map(lambda imgSize, pSize: abs(imgSize - pSize) < OFFSET, imgSizeMM, pStandard)
            if delta.count(True) == len(delta):
                imgSizeMM = pStandard
    try:
        paperFormat = 'A%i' % A.index(tuple(imgSizeMM))
    except ValueError:
        return None
    return paperFormat


def printImage(printer_name, bmp, paperSize=win32con.DMPAPER_A4, tray=None, jobTitle="Untitled"):

    """
    Функция печати изображения по входным параметрам:
    :param printer_name: имя принтера в windows
    :param img:
    :param paperSize:
    :param tray:
     пример использования печати на принтере по-умолчанию на формате А4. Источник бумаги - лоток А4
     printImage(win32print.GetDefaultPrinter(),filename,win32con.DMPAPER_A4,XEROX_A4_TRAY)
    """
    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    hprinter = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
    properties = win32print.GetPrinter(hprinter, 2)
    pr = properties['pDevMode']
    if paperSize == win32con.DMPAPER_A2:
        pr.Orientation = win32con.DMORIENT_LANDSCAPE
    pr.PaperSize = paperSize

    if tray:
        pr.DefaultSource = tray
    properties['pDevMode'] = pr
    try:
        win32print.SetPrinter(hprinter, 2, properties, 0)
    except pywintypes.error, err:
        logger.error(err[2].decode('cp1251').encode('utf-8'))
        sys.exit()
    #
    # You can only write a Device-independent bitmap
    #  directly to a Windows device context; therefore
    #  we need (for ease) to use the Python Imaging
    #  Library to manipulate the image.
    #
    # Create a device context from a named printer
    #  and assess the printable size of the paper.
    #
    gDC = win32gui.CreateDC("WINSPOOL", printer_name, pr)
    hDC = win32ui.CreateDCFromHandle(gDC)
    printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
    printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
    printer_margins = hDC.GetDeviceCaps(PHYSICALOFFSETX), hDC.GetDeviceCaps(PHYSICALOFFSETY)

    #
    # Open the image, rotate it if it's wider than
    #  it is high, and work out how much to multiply
    #  each pixel by to get it as big as possible on
    #  the page without distorting.
    #

    if ((bmp.size[1] < bmp.size[0] and pr.Orientation == win32con.DMORIENT_PORTRAIT) or
                    (bmp.size[0] < bmp.size[1] and pr.Orientation == win32con.DMORIENT_LANDSCAPE)):
        bmp = bmp.rotate(90)

    ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
    scale = min (ratios)

    #
    # Start the print job, and draw the bitmap to
    #  the printer device at the scaled size.
    #
    hDC.StartDoc(jobTitle)
    hDC.StartPage()
    bmp.mode = "L"
    dib = ImageWin.Dib(bmp)
    scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
    x1 = int((printer_size[0] - scaled_width) / 2) - printer_margins[0]
    y1 = int((printer_size[1] - scaled_height) / 2) - printer_margins[1]
    x1 = 0
    y1 = 0
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height
    dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()


def getPrinterByPaperFormat(max_PaperFormat):
    for printer in PRINTERS.items():
        if max_PaperFormat in printer[1]['PaperFormats']:
            return PRINTERS[printer[0]]
    return None


def autoPrintImage(file_name,  max_PaperFormat="A1"):
    try:
        img = Image.open(file_name)
    except IOError:
        logger.error(u"Ошибка при открытии файла %s" % file_name)
        return False
    frm = getImagePaperFormat(img)
    if PAPER_FORMATS.index(frm) > PAPER_FORMATS.index(max_PaperFormat):
        frm = max_PaperFormat
    if frm:
        printer = getPrinterByPaperFormat(frm)
        if not frm in printer['PaperFormats']:
            print u"Выбранный принтер не может распечатать %s. Формат %s не поддерживается принтером %s" % (file_name, frm, printer['Name'])
            raise PaperFormatException(u"Выбранный принтер не может распечатать %s. "
                                           u"Формат %s не поддерживается принтером %s"
                                            % (file_name, frm, printer['Name']))
        tray = None

        if type(printer['PaperFormats']) == dict:
            tray = printer['PaperFormats'][frm]
        curFile = os.path.split(file_name)[1]
        logger.info(u"Печать файла %s на принтере %s (%s)" % (curFile, printer['Name'], frm))
        frm = getattr(win32con, "DMPAPER_" + frm)
        if not __debug__:
            printImage(printer['Name'], img, frm, tray, jobTitle=curFile)
        return True
    else:
        printSize = getImageSizeInMM(img)
        logger.warning(u"Файл %s не был распечатан. Неизвестный формат %sx%s мм" % (file_name, printSize[0], printSize[1]))


if __name__ == '__main__':
    for i in xrange(1, 8):
        autoPrintImage(ur"D:\проекты\Сателит\чертежи\СКИД.301319.052СБ саттелит(%i).tif" % i, max_PaperFormat="A4")
