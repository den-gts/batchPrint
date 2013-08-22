# -*-coding: utf-8 -*-
import os, logging, sys
logger = logging.getLogger("main.misc")
logger.setLevel(logging.DEBUG)


def getFileName(filePath):
    fileNameExt = os.path.basename(filePath)
    return os.path.splitext(fileNameExt)[0]


def getFileList(path, extensions=None, recursive=True):
    osPath = os.path
    if not recursive:
        pathList = os.listdir(path)
        files = map(lambda x: osPath.normpath(osPath.join(path, x)), pathList)
        files = filter(lambda x: osPath.isfile(x), files)
    else:
        files = []
        for curFolder, Folders, filesInFolder in os.walk(path):
            for fileName in filesInFolder:
                files.append(osPath.normpath(osPath.join(curFolder, fileName).replace('\\', '/')))
    if extensions:
        files = filter(lambda x: (osPath.splitext(x)[1].lower() in extensions) and
                                osPath.splitext(x)[1],
                       files)
    return files


def getFSlogger(LoggerName, filePath):
    log = logging.getLogger(LoggerName)
    log.setLevel(logging.DEBUG)
    strH = logging.StreamHandler()
    formatter = logging.Formatter(u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s')
    strH.setFormatter(formatter)
    log.addHandler(strH)
    logFile = logging.FileHandler(filePath)
    logFile.setFormatter(formatter)
    log.addHandler(logFile)
    return log


def checkWritePath(path):
    try:
        open(path + "/dummy", "w")
    except IOError:  # если не удалось записать файл в папку
        return False
    else:
        try:
            os.remove(path + "/dummy")
        except (IOError, WindowsError):
            pass
    return True

if __name__ == "__main__":
    logger.info(__name__)

