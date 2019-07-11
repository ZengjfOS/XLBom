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
key cols        = [1, 2, 5]
keep cols       = [0, 3, 4]
amount col      = 6
parts col        = 7
'''

# excel file title length
titleLength = 0
# keys index for concat compare
keysString = ["Value", "Description", "[Decal]"]
keysIndex = []
# amount column index
amountString = "Qty."
amountCol = 0
# concat parts column index
partsString = "Part(s)"
partsCol = 0
# keep column, just match first cell program find
keepCols = []
currentFilePath = ""


def checkCellEmpty(sheet, row, col):

    if len(str(sheet.row_values(row)[col])) == 0:
        print("\nplease check file: " + currentFilePath + " row <" + str(row + 1) + "> coloum <" + str(col + 1) + "> is Empty?")
        exit(0)

    return sheet.row_values(row)[col]

def analyzeTitle(cols):
    global titleLength
    global partsCol
    global amountCol
    global keysIndex
    global keepCols

    titleLength = 0
    partsCol    = 0
    amountCol   = 0
    keysIndex   = []
    keepCols    = []

    titleLength = len(cols)
    for count in range(0, titleLength):
        print(cols[count])

        if cols[count] == partsString:
            partsCol = count
        elif cols[count] == amountString:
            amountCol = count
        elif cols[count] in keysString:
            keysIndex.append(count)
        else:
            keepCols.append(count)

        count += 1

    print(titleLength, partsCol, amountCol, keysIndex, keepCols)

    if partsCol == 0 or amountCol == 0 or titleLength == 0 or len(keepCols) == 0 or len(keysIndex) < len(keysString):
        print("\ncurrent title is: [ " + ", ".join(col for col in cols) + " ]")
        print("Please check your keys string value: [ " + ", ".join(key for key in keysString) + " ]")
        exit(-1)


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

    currentFilePath = inputFile
    bomXL = xlrd.open_workbook(filename=inputFile)

    outputBomXL = xlwt.Workbook()
    print(bomXL.sheet_names())
    outputBom = outputBomXL.add_sheet(bomXL.sheet_names()[0], cell_overwrite_ok = True)

    bom = bomXL.sheet_by_index(0)
    print(bom.name, bom.nrows, bom.ncols, bom.row_values(0))
    analyzeTitle(bom.row_values(0))

    # copy title
    for index in range(titleLength):
        outputBom.write(0, index, bom.cell(0, index).value)

    # concat key columns(keysIndex) string as key save at Set
    # keySet = set()
    keySet = []
    print("key set -->")
    for row in range(1, bom.nrows):
        # print(bom.row_values(row))

        keyString = ""

        # concat key
        for keyIndex in keysIndex:
            
            if len(keyString) == 0:
                keyString = checkCellEmpty(bom, row, keyIndex)
            else:
                keyString += "|"+ checkCellEmpty(bom, row, keyIndex)

        if keyString not in keySet:
            keySet.append(keyString)
        else: 
            print("key exist in keySet, skip " + keyString + ", row <" + str(row + 1) + ">")

        print(keyString)
    
    print("key set <--")
    # print(keySet)

    # every key in keySet is a line
    outRow = 1
    for key in keySet:
        amount = 0
        parts = "" 
        lineValues = []
        keepColsValues = []

        print("-----------------------row-----------------------")

        # out row start 1, row 0 is title line
        for row in range(1, bom.nrows):
            keyString = ""

            # concat key
            for keyIndex in keysIndex:
                if len(keyString) == 0:
                    keyString = bom.row_values(row)[keyIndex]
                else:
                    keyString += "|"+ bom.row_values(row)[keyIndex]

            # get amount and parts
            if (key == keyString):
                amount += checkCellEmpty(bom, row, amountCol)
                if len(parts) == 0:
                    parts = checkCellEmpty(bom, row, partsCol)
                else:
                    parts += "," + checkCellEmpty(bom, row, partsCol)

                rowKeepColsValues = []
                for col in keepCols:
                    rowKeepColsValues.append(bom.row_values(row)[col])

                keepColsValues.append(rowKeepColsValues)
            
        # add a row value to array
        for col in range(titleLength):
            keysIndexCount = 0
            for colCheck in keysIndex:
                if col == colCheck:
                    lineValues.append(key.split("|")[keysIndexCount])

                keysIndexCount += 1
            
            if col == amountCol:
                lineValues.append(amount)

            if col == partsCol:
                lineValues.append(parts)

            keysIndexCount = 0
            for colCheck in keepCols:
                if col == colCheck:
                    keepColValuesConcat = ""
                    for keepColValues in keepColsValues:
                        if (keysIndexCount == 0):
                            if isinstance(keepColValues[keysIndexCount], int) or isinstance(keepColValues[keysIndexCount], float):
                                keepColValuesConcat += str(int(keepColValues[keysIndexCount])) + ","
                            else:
                                keepColValuesConcat += str(keepColValues[keysIndexCount]) + ","
                        else:
                            keepColValuesConcat = str(keepColValues[keysIndexCount]) + ""

                    if (keysIndexCount == 0 and len(keepColValuesConcat) > 0):
                        lineValues.append(keepColValuesConcat[0:-1])
                    else:
                        lineValues.append(keepColValuesConcat)

                keysIndexCount += 1

        # write to output bom file
        for col in range(titleLength):
            outputBom.write(outRow, col, lineValues[col])

        print(lineValues)
        print(keepColsValues)

        outRow += 1

    outputBomXL.save(outputFile)

if __name__ == '__main__':
    files = getCSVFiles()
    for file in files:
        bom("inputs/" + file, "outputs/" + os.path.splitext(file)[0] + "_" + datetime.now().strftime("%Y%m%d%H%M%S") + ".xls")