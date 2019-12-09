#!/usr/bin/python
# -*- coding: utf-8 -*-

from os import walk
import os
import xlrd
import xlwt
from datetime import datetime

'''
|<-----------columns----------->|
+-------------------------------+
| 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
+-------------------------------+

title length    = 8
key cols        = [6]
keep cols       = [0, 3, 4]
'''

# excel file title length
titleLength = 0
# keys index for concat compare
keysString = ["Manufacturer"]
keysIndex = []
# keep column, just match first cell program find
keepColsString = ["Item", "Value", "Description", "Manufacturer", "[Decal]", "尚缺"]
keepCols = []
currentFilePath = ""

# debug info 
debug = True


def checkCellEmpty(sheet, row, col):

    if len(str(sheet.row_values(row)[col])) == 0:
        print("\nplease check file: " + currentFilePath + " row <" + str(row + 1) + "> coloum <" + str(col + 1) + "> is Empty?")
        exit(0)

    return sheet.row_values(row)[col]

def analyzeTitle(cols):
    global titleLength
    global keysIndex
    global keysString
    global keepColsString
    global keepCols

    keysIndex   = []
    keepCols = []

    # get length of title
    titleLength = len(cols)

    # analyze the title array
    for count in range(0, titleLength):
        # show current col
        # print(cols[count])

        if cols[count] in keysString:     # get "amount" column number
            keysIndex.append(count)
        
        if cols[count] in keepColsString:
            keepCols.append(count)

    # check title index
    if len(keysIndex) < len(keysString) or len(keepCols) == 0 or len(keepCols) != len(keepColsString):
        print("\ncurrent title is: [ " + ", ".join(col for col in cols) + " ]")
        print("Please check your keys string value: [ " + ", ".join(key for key in keysString) + " ]")
        exit(-1)

    print("\nKey Column: [ " + ", ".join(str(key) for key in keysIndex) + " ]")
    print("\nColumns: [ " + ", ".join(str(key) for key in keepCols) + " ]")


def getCSVFiles():

    f = []
    for (dirpath, dirnames, filenames) in walk("inputs") :
        f.extend(filenames)

    for file in f:
        print(file)

    if len(f) == 0:
        print("\nPlease check your input file, there is empty.")
        exit(-2)

    return f

def bom(inputFile, outputFile):

    print(inputFile)
    print(outputFile)

    global currentFilePath

    # open input BOM file
    currentFilePath = inputFile
    bomXL = xlrd.open_workbook(filename=inputFile)

    # print sheet array in input BOM file
    print(bomXL.sheet_names())

    # get input file's sheet[0]
    bom = bomXL.sheet_by_index(0)

    # show sheet infos, like: sheet name, sheet rows, sheet cols, first row of sheet(sheet title)
    print(bom.name, bom.nrows, bom.ncols, bom.row_values(0))

    # analyze sheet title
    analyzeTitle(bom.row_values(0))

    # concat key columns(keysIndex) string as key save at Set
    # keySet = set()
    keySet = []
    print("key set -->")
    for row in range(1, bom.nrows):
        # print(bom.row_values(row))

        keyString = (bom.row_values(row))[keysIndex[0]]

        if len(keyString.strip()) == 0:
            print("\nError: please check excel row " + str(row) + "is no empty")
            exit(-1)

        if keyString not in keySet:
            keySet.append(keyString)
        else: 
            # print("key exist in keySet, skip " + keyString + ", row <" + str(row + 1) + ">")
            pass

        # print(keyString)
    
    print("key set <--")
    print(keySet)

    # every key in keySet is a line
    outRow = 1
    for key in keySet:
        # open BOM output file
        outputBomXL = xlwt.Workbook()
        # add same sheet[0] name to output file from input file sheet[0] name, we just analyze sheet[0]
        outputBom = outputBomXL.add_sheet(bomXL.sheet_names()[0], cell_overwrite_ok = True)

        '''
        # copy title
        for index in range(titleLength):
            outputBom.write(0, index, bom.cell(0, index).value)
        '''
        for index in range(len(keepCols)):
            outputBom.write(0, index, bom.cell(0, keepCols[index]).value)

        rowCount = 1
        for row in range(1, bom.nrows):
            if (bom.row_values(row))[keysIndex[0]] == key :
                for index in range(len(keepCols)):
                    outputBom.write(rowCount, index, bom.row_values(row)[keepCols[index]])

                rowCount += 1

        outputBomXL.save(outputFile + key + "_" + datetime.now().strftime("%Y%m%d%H%M%S") + ".xls")


if __name__ == '__main__':
    files = getCSVFiles()
    for file in files:
        # bom("inputs/" + file, "outputs/" + os.path.splitext(file)[0])
        bom("inputs/" + file, "outputs/")