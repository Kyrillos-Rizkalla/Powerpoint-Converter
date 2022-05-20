import sys
import os
import re
import shutil
import tempfile
import traceback
import win32com.client as win32
from pathlib import Path


def show_exception_and_exit(exc_type, exc_value, tb):
    traceback.print_exception(exc_type, exc_value, tb)
    input("Press key to exit.")
    sys.exit(-1)


class Jes_Functions:
    def __init__(self, dataDir):
        root = dataDir.decode()
        self.root = root

    def done_Path(self):
        donedir = self.root + '\\_Done_'
        if not os.path.exists(donedir):
            # Create a new directory because it does not exist
            os.makedirs(donedir)

        self.done_root = donedir
        print()
        print('All successfully converted files will be moved to: ')
        print(donedir)
        print()

    def get_allPath(self, suffix_s):
        files = []
        for root, dirs, _files in os.walk(self.root):
            dirs[:] = [d for d in dirs if not re.match('_Done_', d)]
            for file in _files:
                file = os.path.abspath(os.path.join(root, file))
                for suffix in suffix_s:
                    if Path(file).suffix == suffix:
                        files.append(str(file))
        return files

    # @staticmethod
    def ppt_converter(self, formatType):
        if formatType == 'pptx':
            source_Paths = self.get_allPath(['.ppt'])
            if not source_Paths:
                print('No .PPT files found in current directory.')
            else:
                pptx_Paths = [source_Path.replace(
                    '.ppt', '.pptx') for source_Path in source_Paths]

        elif formatType == 'pdf':
            source_Paths = self.get_allPath(['.pptx', '.ppt'])
            if not source_Paths:
                print('No .PPT or .PPTx files found in current directory.')
            else:
                pdf_Paths = [source_Path.replace('.pptx', '.pdf').replace(
                    '.ppt', '.pdf') for source_Path in source_Paths]

        if source_Paths:
            totalfiles = len(source_Paths)
            self.done_Path()

            for (index, source_Path) in enumerate(source_Paths):
                ppt = win32.gencache.EnsureDispatch("PowerPoint.Application")
                pres = ppt.Presentations.Open(source_Path, True, False, False)

                if formatType == 'pptx':
                    pres.SaveAs(pptx_Paths[index])
                elif formatType == 'pdf':
                    # formatType = 32 for ppt to pdf
                    pres.SaveAs(pdf_Paths[index], 32)

                pres.Close()
                ppt.Quit()

                os.rename(source_Path, source_Path.replace(
                    self.root, self.done_root))

                print(source_Path)
                print(str(round((index + 1)/totalfiles*100)).zfill(2) +
                      '% ('+str(index + 1).zfill(len(str(totalfiles))) +
                      ' of ' + str(totalfiles) +
                      ') >>> File converted successfully.')
                print()


sys.excepthook = show_exception_and_exit

texp_path = tempfile.gettempdir() + r'\gen_py'

if os.path.isdir(texp_path):
    shutil.rmtree(texp_path)

path = os.getcwdb()

print('Working path: ' + path.decode())

print(), print()
input("Press Enter to continue...")

cnvrt = Jes_Functions(path)

print(), print()
print("f: to convert (ppt and pptx) files to pdf")
print("x: to convert ppt files to pptx")
print("enter: to exit.")
print()

working = True
while working:

    m = input("Press select:")
    print(m)
    if m in ['F', 'f']:
        print('Converting .(ppt and pptx) files to pdf')
        print(), print()
        cnvrt.ppt_converter('pdf')
        working = False

    elif m in ['X', 'x']:
        print('Converting .ppt files to pptx')
        print(), print()
        cnvrt.ppt_converter('pptx')
        working = False

    elif m == '':
        working = False

    else:
        print()
        print(r"Please make sure you press 'f' or 'x' or 'enter'...")
        print(), print()

if os.path.isdir(texp_path):
    shutil.rmtree(texp_path)

print(), print('Thanks for using Powerpoint Converter...'), print()
print('++++++++++++++++++++++++++++++++++++++++++++')

print(), print()
input("Press Enter to Exit...")
