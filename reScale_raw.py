import fitz
import os
import subprocess
import shutil
import glob
import re
import time

def preProcessURAR(pdfFile):
    fileType = re.findall('\d{4}', pdfFile)[0]
    cwdL = pdfFile.split(fileType)
    cwd = cwdL[0]
    pdfFile = f'{fileType}{cwdL[1]}'.split('.')[0]
    # Top
    top = 0
    # Bottom
    bot = 0
    # absoluteOffset Left
    aosL = 0
    # absoluteOffset Right
    aosR = 0
    return pdfFile, top, bot, aosL, aosR, cwd

def cleanUpPDFs(pdfFile, cwd):
    # DELETE cropped & RENAME reScaled
    os.remove(f'{cwd}\\{pdfFile}.pdf')
    shutil.move(f'{cwd}\\{pdfFile}_reScaled.pdf', f'{cwd}\\{pdfFile}.pdf')
    print('Replaced Original File')


def countDirectories(CWD):
    directoryList = []
    for directory in os.listdir(CWD):
        try:
            if os.listdir("{}/{}".format(CWD, directory)):
                directoryList.append("{}/{}".format(CWD, directory))
                # print("{}/{}".format(CWD, directory))
        except:
            continue
    return directoryList

def countFilesInEachDirectory(directoryList):
    allFiles = []
    for directory in directoryList:
        fileList = glob.glob(directory+'/'+'*.pdf')
        allFiles.append(fileList)
    return allFiles


# MASTER FUNCTION
def reFormPDFs():
    directoryList = countDirectories('E:\TEMP') # ROOT DIRECTORY OF EXECUTION: WILL LOOK IN SUB-FOLDERS HERE
    allFiles = countFilesInEachDirectory(directoryList)
    for directory in allFiles:
        for file in directory:
            regex_match = False
            pdfFile, top, bot, aosL, aosR, cwd = preProcessURAR(file)
            ############################################################################################
            # ADD ANY UNIQUE PART OF A FILE NAME TO THIS LIST TO ADD PDF FILES YOU WANT TO BE RESCALED #
            ############################################################################################
            regex_list = ['_C\d-\d_', 'CERT', 'COST', 'MC']
            for pattern in regex_list:
                if re.search(pattern, file):
                    print(file)
                    regex_match = True

            # UNCOMMENT THIS LINE IF YOU WANT ALL PDF FILES TO BE RE_SCALED WITH NO CONDITIONS.
            # regex_match = True

            if regex_match:
                # RE-SCALE PAGE
                reScaleCommand = ['C:\\gs\\gs9_27\\bin\\gswin64.exe', '-sDEVICE=pdfwrite', '-o', f'{pdfFile}_reScaled.pdf', '-CompatibilityLevel=1.4', '-sPAPERSIZE=legal', '-dFIXEDMEDIA', '-dPDFFitPage', f'{pdfFile}.pdf']
                reScaledPDF = subprocess.Popen(reScaleCommand, text=True, shell=True, cwd=f'{cwd}')
                reScaledPDF.wait()
                print('Re-Scaled PDF file.')
                cleanUpPDFs(pdfFile, cwd)

# POINT OF CODE EXECUTION
reFormPDFs()