import xlrd
import xlwt

def boom():

    boomXL = xlrd.open_workbook(filename="module.xlsx")

    outputBoomXL = xlwt.Workbook()
    outputBoom = outputBoomXL.add_sheet('boom', cell_overwrite_ok = True)

    print(boomXL.sheet_names())

    boom = boomXL.sheet_by_index(0)
    print(boom.name, boom.nrows, boom.ncols)

    outputBoom.write(0, 0, boom.cell(0, 0).value)
    outputBoom.write(0, 1, boom.cell(0, 1).value)
    outputBoom.write(0, 2, boom.cell(0, 2).value)
    outputBoom.write(0, 3, boom.cell(0, 3).value)

    keySet = set()
    for index in range(1, boom.nrows):
        print(boom.row_values(index))
        keyString = boom.row_values(index)[0] + "|"+ boom.row_values(index)[1]
        keySet.add(keyString)
    
    print(keySet)

    i = 1
    for key in keySet:
        amount = 0
        info = "" 

        for index in range(1, boom.nrows):
            keyString = boom.row_values(index)[0] + "|" + boom.row_values(index)[1]
            if (key == keyString):
                amount += boom.row_values(index)[2]
                info += "," + boom.row_values(index)[3]

        outputBoom.write(i, 0, key.split("|")[0])
        outputBoom.write(i, 1, key.split("|")[1])
        outputBoom.write(i, 2, amount)
        outputBoom.write(i, 3, info)

        i += 1

    outputBoomXL.save('output.xls')

if __name__ == '__main__':
    boom()