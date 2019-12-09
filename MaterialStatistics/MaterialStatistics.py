#!/usr/bin/python
# -*- coding: utf-8 -*-

from os import walk
import os
from datetime import datetime
import pandas
import codecs

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

if __name__ == '__main__':
    files = getCSVFiles()
    for file in files:
        xd = pandas.ExcelFile("inputs/" + file)
        df = xd.parse(xd.sheet_names[0], header=None, keep_default_na=False)
        with codecs.open("outputs/" + os.path.splitext(file)[0] + "_" + datetime.now().strftime("%Y%m%d%H%M%S") + ".html", 'w+',"utf-8") as html_file:
            html_file.write(df.to_html(header = True,index = False))
