import fitz
import os
import subprocess
import shutil
import glob
import re
# import pdfplumber
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
# PRE-CROP TOP AND BOTTOM LOGIC 
#               absolutePreCrop4 for C1-C3
#                       Bot     Top                      Left       Right
# TOTAL LETTER 1-3: 0   36 	0   37      absoluteOffset4: 7.0    0	0.6 0
# TOTAL LEGAL  1-3: 0   74 	0   59      absoluteOffset4: 8.5    0	0   0
# WinTOTAL     1-3: 0   69	0   59.5    absoluteOffset4: 8.5    0	0   0
# ClickFORMS   1-3: 0   23	0   59.5    absoluteOffset4: 15.0   0   0   0
# ACI          1-3: 0   55	0   64      absoluteOffset4: 8.0    0   0   0
# FNC          1-3: 0   33	0   37      absoluteOffset4: 8.0    0   0   0
#               absolutePreCrop4 for C4+
# TOTAL LETTER 4+:  0   42  0   36.5    absoluteOffset4: 6.5    0   0   0
# TOTAL LEGAL  4+:  0   74	0   59      absoluteOffset4: 8.5    0   0   0
# WinTOTAL     4+:  0   69 	0   59.5    absoluteOffset4: 8.0    0   0   0
# ClickFORMS   4+:  0   71	0   144	    absoluteOffset4: 15.0   0   0   0
# ACI          4+:  0   55	0   64	    absoluteOffset4: 8.0    0   0   0
# FNC          4+:  0   33	0   37	    absoluteOffset4: 8.0    0   0   0
#
#
#               absolutePreCrop4 for CERT
# TOTAL LETTER 	 :	0	37.5 0	36.5	absoluteOffset4: 0		0	0	0
# TOTAL LEGAL    :	0	74	 0	58		absoluteOffset4: 0		0	0	0
# WinTOTAL       :	0	71	 0	59.2	absoluteOffset4: 0		0	0	0
# ClickFORMS     :	0	35	 0	47.1	absoluteOffset4: 0		0	0.8	0
# ACI            :	0	44.2 0	75.5	absoluteOffset4: 0		0	0	0
# FNC            :	0	32	 0	39		absoluteOffset4: 0		0	0	0
#
#				absolutePreCrop4 for COST
# TOTAL LETTER   :  0	37.5 0	36.5	absoluteOffset4: 7.0    0	0.6 0
# TOTAL LEGAL	 :	0   74 	 0  59      absoluteOffset4: 8.5    0	0   0
# WinTOTAL   	 : 	0   69 	 0  59.5    absoluteOffset4: 8.0    0   0   0
# ClickFORMS     :	0	35.8 0	59.4	absoluteOffset4: 13.0   0   0   0
# ACI            :  0   55	 0  64      absoluteOffset4: 8.0    0   0   0
# FNC            :  0   33	 0  37      absoluteOffset4: 8.0    0   0   0
#
#				absolutePreCrop4 for MC
# TOTAL LETTER   :	0	35	 0	30 		absoluteOffset4: 7.0    0	0.2 0
# TOTAL LEGAL    :	0	72	 0	51		absoluteOffset4: 8.5    0	0   0
# WinTOTAL       :	0	71	 0	52		absoluteOffset4: 8      0	0   0
# ClickFORMS     :	0	35	 0	47		absoluteOffset4: 14     0	0   0
# ACI            :	0	39	 0	64		absoluteOffset4: 8.0	0	0	0
# FNC            :  0   33	0   37      absoluteOffset4: 8.0    0   0   0
    
# COMPS
def comps(pdfFile, top, bot, aosL, aosR, cwd):
    if 'TOTAL' in pdfFile and 'LETTER' in pdfFile:
        top = 37
        bot = 36
        aosL = 7
        aosR = 0.6
        if not '_C1-3_' in pdfFile:
            top = 36.5
            bot = 42
            aosL = 6.5
    elif 'TOTAL' in pdfFile and 'LEGAL' in pdfFile:
        top = 59
        bot = 74
        aosL = 8.5
        if not '_C1-3_' in pdfFile:
            top = 36.5
            bot = 42
            aosL = 6.5
    elif 'WinTOTAL' in pdfFile:
        top = 59.5
        bot = 69
        aosL = 8.5
        if not '_C1-3_' in pdfFile:
            aosL = 8
    elif 'ClickFORMS' in pdfFile:
        top = 59.5
        bot = 23
        aosL = 15
        if not '_C1-3_' in pdfFile:
            top = 144
            bot = 71
    elif 'ACI' in pdfFile:
        top = 64
        bot = 55
        aosL = 8
    elif 'FNC' in pdfFile:
        top = 37
        bot = 33
        aosL = 8
    else:
        print('<UNKNOWN FILE> No Cropping actions were Taken')
    return pdfFile, top, bot, aosL, aosR, cwd 

# CERT
def cert(pdfFile, top, bot, aosL, aosR, cwd):
    if 'TOTAL' in pdfFile and 'LETTER' in pdfFile:
        top = 36.5
        bot = 37.5
    elif 'TOTAL' in pdfFile and 'LEGAL' in pdfFile:
        top = 58
        bot = 74
    elif 'WinTOTAL' in pdfFile:
        top = 59.2
        bot = 71
    elif 'ClickFORMS' in pdfFile:
        top = 47.1
        bot = 35
        aosR = 0.8
    elif 'ACI' in pdfFile:
        top = 75.5
        bot = 44.2
    elif 'FNC' in pdfFile:
        top = 39
        bot = 32
    return pdfFile, top, bot, aosL, aosR, cwd

# COST
def cost(pdfFile, top, bot, aosL, aosR, cwd):
    if 'TOTAL' in pdfFile and 'LETTER' in pdfFile:
        top = 36.5
        bot = 37.5
        aosL = 7
        aosR = 0.6
    elif 'TOTAL' in pdfFile and 'LEGAL' in pdfFile:
        top = 59
        bot = 74
        aosL = 8.5
    elif 'WinTOTAL' in pdfFile:
        top = 59.5
        bot = 69
        aosL = 8
    elif 'ClickFORMS' in pdfFile:
        top = 59.4
        bot = 35.8
        aosL = 13
    elif 'ACI' in pdfFile:
        top = 64
        bot = 55
        aosL = 8
    elif 'FNC' in pdfFile:
        top = 37
        bot = 33
        aosL = 8
    return pdfFile, top, bot, aosL, aosR, cwd

# MC
def mc(pdfFile, top, bot, aosL, aosR, cwd):
    if 'TOTAL' in pdfFile and 'LETTER' in pdfFile:
        top = 30
        bot = 35
        aosL = 7
        aosR = 0.2
    elif 'TOTAL' in pdfFile and 'LEGAL' in pdfFile:
        top = 51
        bot = 72
        aosL = 8.5
    elif 'WinTOTAL' in pdfFile:
        top = 52
        bot = 71
        aosL = 8
    elif 'ClickFORMS' in pdfFile:
        top = 47
        bot = 35
        aosL = 14
    elif 'ACI' in pdfFile:
        top = 64
        bot = 39
        aosL = 8
    elif 'FNC' in pdfFile:
        top = 37
        bot = 33
        aosL = 8
    return pdfFile, top, bot, aosL, aosR, cwd

def cropPDFs(pdfFile, top, bot, aosL, aosR, cwd):
    # CROP MARGINS
    cropCommand = ['pdf-crop-margins','-p','0', '--absolutePreCrop4','0',f'{bot}','0',f'{top}', '--absoluteOffset4', f'{aosL}','0',f'{aosR}','0', rf'{cwd}{pdfFile}.pdf']
    cropPDF = subprocess.Popen(cropCommand, text=True, shell=True, cwd=f'{cwd}')
    cropPDF.wait()
    print('Cropped to [0%] margins')
    raw_file_name = pdfFile.split("\\")[1]
    if f'{raw_file_name}_cropped.pdf' in os.listdir(cwd):
        print("moving", f'{cwd}{raw_file_name}_cropped.pdf', "back to", f'{cwd}\\{pdfFile}_cropped.pdf')
        shutil.move(f'{cwd}{raw_file_name}_cropped.pdf', f'{cwd}\\{pdfFile}_cropped.pdf')
    # exit()
    # RE-SCALE PAGE
    reScaleCommand = ['C:\\gs\\gs9_27\\bin\\gswin64.exe', '-sDEVICE=pdfwrite', '-o', f'{pdfFile}_reScaled.pdf', '-CompatibilityLevel=1.4', '-sPAPERSIZE=legal', '-dFIXEDMEDIA', '-dPDFFitPage', f'{cwd}{pdfFile}_cropped.pdf']
    reScaledPDF = subprocess.Popen(reScaleCommand, text=True, shell=True, cwd=f'{cwd}')
    reScaledPDF.wait()
    print('Re-Scaled PDF file.')
    # CHECK PAGE SIZES
    # with pdfplumber.open(f"{cwd}\\{pdfFile}_reScaled.pdf") as pdf:
    #     pages = pdf.pages[0]
    #     print(f'Width: {pages.width}')
    #     print(f'Height: {pages.height}')

def cleanUpPDFs(pdfFile, cwd):
    # DELETE cropped & RENAME reScaled
    os.remove(f'{cwd}\\{pdfFile}_cropped.pdf')
    shutil.move(f'{cwd}\\{pdfFile}_reScaled.pdf', f'{cwd}\\{pdfFile}.pdf')
    print('Replaced Original File')


def countDirectories(CWD):
    directoryList = []
    for directory in os.listdir(CWD):
        print(directory)
        print("...")
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

def reFormPDFs():
    directoryList = countDirectories('E:\TEMP')
    allFiles = countFilesInEachDirectory(directoryList)
    print(allFiles)
    for directory in allFiles:
        for file in directory:
            file = rf'{file}'
            # Reset default values to 0
            
            pdfFile, top, bot, aosL, aosR, cwd = preProcessURAR(file)
            # COMPS
            if re.search('_C\d-\d_', file):                                         #TODO C1-4_ACI is not cropping correctly IDK why... but I will sleep now :)
                pdfFile, top, bot, aosL, aosR, cwd = comps(pdfFile, top, bot, aosL, aosR, cwd)
            # CERT
            elif re.search('CERT', file):
                pdfFile, top, bot, aosL, aosR, cwd = cert(pdfFile, top, bot, aosL, aosR, cwd)
            # COST
            elif re.search('COST', file):
                pdfFile, top, bot, aosL, aosR, cwd = cost(pdfFile, top, bot, aosL, aosR, cwd)
            # MC
            elif re.search('MC', file):
                pdfFile, top, bot, aosL, aosR, cwd = mc(pdfFile, top, bot, aosL, aosR, cwd)
            else:
                continue
            print(">>>>>>>>>>",f'{cwd}{pdfFile}')
            cropPDFs(pdfFile, top, bot, aosL, aosR, cwd)
            cleanUpPDFs(pdfFile, cwd)
reFormPDFs()
