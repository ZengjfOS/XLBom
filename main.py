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
outputSheetName = "boom"

def boom():

    boomXL = xlrd.open_workbook(filename=inputXLFileName)

    outputBoomXL = xlwt.Workbook()
    outputBoom = outputBoomXL.add_sheet(outputSheetName, cell_overwrite_ok = True)

    print(boomXL.sheet_names())

    boom = boomXL.sheet_by_index(0)
    print(boom.name, boom.nrows, boom.ncols)

    # copy title
    for index in range(titleLength):
        outputBoom.write(0, index, boom.cell(0, index).value)

    # concat key columns(keysIndex) string as key save at Set
    keySet = set()
    for row in range(1, boom.nrows):
        print(boom.row_values(row))

        keyString = ""

        # concat key
        for keyIndex in keysIndex:
            if len(keyString) == 0:
                keyString = boom.row_values(row)[keyIndex]
            else:
                keyString += "|"+ boom.row_values(row)[keyIndex]

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
        for row in range(1, boom.nrows):
            keyString = ""

            # concat key
            for keyIndex in keysIndex:
                if len(keyString) == 0:
                    keyString = boom.row_values(row)[keyIndex]
                else:
                    keyString += "|"+ boom.row_values(row)[keyIndex]

            # get amount and info
            if (key == keyString):
                amount += boom.row_values(row)[amountCol]
                if len(info) == 0:
                    info = boom.row_values(row)[infoCol]
                else:
                    info += "," + boom.row_values(row)[infoCol]

            for col in keepCols:
                keepColsValues.append(boom.row_values(row)[col])
            
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

        # write to output boom file
        for col in range(titleLength):
            outputBoom.write(outRow, col, lineValues[col])

        outRow += 1

    outputBoomXL.save(outputXLFileName)

if __name__ == '__main__':
    boom()