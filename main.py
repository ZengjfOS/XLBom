#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlrd
import xlwt

# excel file title length
titleLength = 4
# keys index for concat
keysIndex = [0, 1]
# amount column index
amountCol = 2
# info column index
infoCol = 3
# keep column, just match last cell program find
keepCols = []
# input excel file name
inputXLFileName = "module.xlsx"
# output excel file name
outputXLFileName = "output.xls"
# output sheet name
outputSheetName = "bom"

def bom():

    bomXL = xlrd.open_workbook(filename=inputXLFileName)

    outputBomXL = xlwt.Workbook()
    outputBom = outputBomXL.add_sheet(outputSheetName, cell_overwrite_ok = True)

    print(bomXL.sheet_names())

    bom = bomXL.sheet_by_index(0)
    print(bom.name, bom.nrows, bom.ncols)

    # copy title
    for index in range(titleLength):
        outputBom.write(0, index, bom.cell(0, index).value)

    # concat key columns(keysIndex) string as key save at Set
    keySet = set()
    for row in range(1, bom.nrows):
        print(bom.row_values(row))

        keyString = ""

        # concat key
        for keyIndex in keysIndex:
            if len(keyString) == 0:
                keyString = bom.row_values(row)[keyIndex]
            else:
                keyString += "|"+ bom.row_values(row)[keyIndex]

        keySet.add(keyString)
    
    print(keySet)

    # every key in keySet is a line
    outRow = 1
    for key in keySet:
        amount = 0
        info = "" 
        lineValues = []
        keepColsValues = []

        # out row start 1, row 0 is title line
        for row in range(1, bom.nrows):
            keyString = ""

            # concat key
            for keyIndex in keysIndex:
                if len(keyString) == 0:
                    keyString = bom.row_values(row)[keyIndex]
                else:
                    keyString += "|"+ bom.row_values(row)[keyIndex]

            # get amount and info
            if (key == keyString):
                amount += bom.row_values(row)[amountCol]
                if len(info) == 0:
                    info = bom.row_values(row)[infoCol]
                else:
                    info += "," + bom.row_values(row)[infoCol]

            for col in keepCols:
                keepColsValues.append(bom.row_values(row)[col])
            
        # add a row value to array
        for col in range(titleLength):
            keysIndexCount = 0
            for colCheck in keysIndex:
                if col == colCheck:
                    lineValues.append(key.split("|")[keysIndexCount])

                keysIndexCount += 1
            
            if col == amountCol:
                lineValues.append(amount)

            if col == infoCol:
                lineValues.append(info)

            keysIndexCount = 0
            for colCheck in keepCols:
                if col == colCheck:
                    lineValues.append(keepColsValues[keysIndexCount])

                keysIndexCount += 1

        # write to output bom file
        for col in range(titleLength):
            outputBom.write(outRow, col, lineValues[col])

        outRow += 1

    outputBomXL.save(outputXLFileName)

if __name__ == '__main__':
    bom()