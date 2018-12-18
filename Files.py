import glob
import os.path
import mmap
import re

infiles = r"E:\GitHub\MailMiner\Sample Inputs\Wellesbourne Weather\MMS_Daily_Wellesbourne_20160928.csv"

def FileGenerator(infiles,regex):
    for fullpath in glob.iglob(infiles):
        filename = os.path.basename(fullpath)
        regexmatch = regex.fullmatch(filename)
        if regexmatch:
            with open(fullpath, "rb") as file:
                with mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as bytedata:
                    yield {b"filename":filename, "bytedata":bytedata, b"regexmatch":regexmatch}

