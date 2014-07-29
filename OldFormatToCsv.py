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

setting = {
    "bitstream_filename_regex":[
        '.+-(\d+)',
        '(\d+)'
    ], # the second group of regex match result is the priority of bitstream

    'xls_col2csv_col_preReplace':[':','.'],
    'xls_col2csv_col':[
        [r'(.+)=(.+)',r'dc.\1[\2]'],
        [r'(.+)',r'dc.\1[zh_TW]']
    ],

    'dataDelimiter':';',
    "multiField":[
        "dc.contributor",
        "dc.contributor.advisor",
        "dc.contributor.author",
        "dc.contributor.editor",
        "dc.contributor.illustrator",
        "dc.contributor.other",
        "dc.contributor.cataloger",
        "dc.creator",
        "dc.subject",
        "dc.subject.classification",
        "dc.subject.ddc",
        "dc.subject.lcc",
        "dc.subject.lcsh",
        "dc.subject.mesh",
        "dc.subject.other",
    ],
}

import sys,csv,os,re
from xlrd import open_workbook

def parseXls(xls_file):
    wb = open_workbook(xls_file)

    # print('# Xls [{0}] ======='.format(xls_file))

    items=[]

    for s in wb.sheets():
        # print('# Sheet [{0}] ======='.format(s.name))

        # print('## this sheet has {0} cols and {1} rows.'.format(s.ncols,s.nrows))
        col_name=[]
        ncols=range(s.ncols)

        for col in ncols:
            #print("...",s.cell(0,col).value)
            col_name.append(s.cell(0,col).value)

        for row in range(1,s.nrows):
            item=[]
            # print('### item {0} ======='.format(row))
            for col in ncols:
                item.append(s.cell(row,col).value)
                # print('[{0}] : [{1}]'.format(col_name[col],s.cell(row,col).value))
            items.append(item)
    print()
    return {'items':items,'col_name':col_name}

def file_sniff(target_path):
    if not os.path.isdir(target_path):
        raise BaseException("Not a folder!")

    fs = os.listdir(target_path)
    bitstream_filename_regex_index = 0
    bitstream_filename_regex_index_locked = False
    xls_file = None
    bs_reg = re.compile(setting['bitstream_filename_regex'][bitstream_filename_regex_index])
    bitsteams = []

    # print("files in this path:",fs)
    for f in fs:
        f_detail = os.path.splitext(f)

        #print(fileExtension[1])
        
        if f_detail[1].lower() in ['.xls','.xlsx']:
            xls_file = f
        else:
            reg_match = bs_reg.match(f_detail[0])
            while not reg_match:
                if bitstream_filename_regex_index_locked:
                    raise BaseException("bitstream filename format not the same!")
                bs_reg = re.compile(setting['bitstream_filename_regex'][bitstream_filename_regex_index])
                bitstream_filename_regex_index+=1
                reg_match = bs_reg.match(f_detail[0])

            bitstream_filename_regex_index_locked = True
            bitsteams.append([int(reg_match.group(1)),f])

    bitsteams.sort( key = lambda ele: ele[0] )
    bitsteams = [ bs[1] for bs in bitsteams ]

    return {'xls_file':xls_file,'bitsteams':bitsteams}

def convert(parsedXls,bitsteams):
    print(parsedXls['col_name'])
    col_name = []
    multiField_k = []
    k = 0
    for col in parsedXls['col_name']:
        col_name_tmp = col.replace(setting['xls_col2csv_col_preReplace'][0],setting['xls_col2csv_col_preReplace'][1])


        for xls_col2csv_col in setting['xls_col2csv_col']:
            if re.match(xls_col2csv_col[0],col_name_tmp):
                col_name_tmp = re.sub(xls_col2csv_col[0],xls_col2csv_col[1],col_name_tmp)
                col_name.append(col_name_tmp)
                if col_name_tmp in setting['multiField']:
                    multiField_k.append(k) # this place has some problem to solve...
                break
        k+=1
    
    print(col_name)
    print(multiField_k)


def oldFormat2Csv(source_path,output_path):
    file_sniff_result = file_sniff(source_path)
    parsedXls = parseXls(os.path.join(source_path,file_sniff_result['xls_file']))
    convert(parsedXls,file_sniff_result['bitsteams'])

print(oldFormat2Csv(sys.argv[1],sys.argv[2]))
