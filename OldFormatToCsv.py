#!/usr/bin/python

# xls_file="metadata_01.xls"

# print("Hello, Python!")
# from mmap import mmap,ACCESS_READ
# from xlrd import open_workbook
# print(open_workbook(xls_file))
# with open(xls_file,'rb') as f:
#     print(open_workbook(
#         file_contents=mmap(f.fileno(),0,access=ACCESS_READ)
#         ))
# aString = open(xls_file,'rb').read()
# print(open_workbook(file_contents=aString))

import sys,csv
from xlrd import open_workbook

def parseXls(xls_file):
    wb = open_workbook(xls_file)

    print('# Xls [{0}] ======='.format(xls_file))

    items=[]

    for s in wb.sheets():
        print('# Sheet [{0}] ======='.format(s.name))

        print('## this sheet has {0} cols and {1} rows.'.format(s.ncols,s.nrows))
        col_name=[]
        ncols=range(s.ncols)

        for col in ncols:
            #print("...",s.cell(0,col).value)
            col_name.append(s.cell(0,col).value)

        for row in range(1,s.nrows):
            item={}
            print('### item {0} ======='.format(row))
            for col in ncols:
                item[col_name[col]]=s.cell(row,col).value
                print('[{0}] : [{1}]'.format(col_name[col],s.cell(row,col).value))
            items.append(item)
    print()
    return items


print(parseXls(sys.argv[1]))
