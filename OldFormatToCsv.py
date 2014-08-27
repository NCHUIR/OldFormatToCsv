#!/usr/bin/python

import sys,csv,os,re
from xlrd import open_workbook

class OldFormatToCsv:

    setting = {
        "bitstream_filename_regex":[
            ['dash(-)','^.+-(\d+)$'],
            ['pure_digits','^(\d+)$']
        ], # the second group (1) of regex match result is the priority of bitstream

        'xls_ext':['xls','xlsx'],

        'bitstream_ext':['pdf'],

        'xls_col2csv_col_preReplace':[':','.'],
        'xls_col2csv_col':[
            [r'(.+)=(.+)',r'dc.\1[\2]'],
            [r'(.+)',r'dc.\1[zh_TW]']
        ], # the second group (1) of regex match result is the name for matching multiField and ignored_col

        'oldDelimiter2New':[';','||'],
        "multiField":[
            "contributor",
            "contributor.advisor",
            "contributor.author",
            "contributor.editor",
            "contributor.illustrator",
            "contributor.other",
            "contributor.cataloger",
            "creator",
            "subject",
            "subject.classification",
            "subject.ddc",
            "subject.lcc",
            "subject.lcsh",
            "subject.mesh",
            "subject.other",
        ],
        'ignored_col':[
            'sys_filename',
            'sys_replace'
        ],

        'pattern_to_remove':['\n'],

        'bitstream_name_col':'contents',
    }

    Meta = {'items':[],'col_name':[]}

    xls_file = None
    bitsteams = []
    bitstream_name_format = None
    ignored_files = []

    def __init__(self,input_setting = {}):
        self.setting.update(input_setting)

    @staticmethod
    def remove_ptn(str,ptn_to_remove):
        for ptn in ptn_to_remove:
            str = str.replace(ptn,'')
        return str

    @staticmethod
    def parseXls(xls_file):
        wb = open_workbook(xls_file)

        items = []

        for s in wb.sheets():
            col_name = []
            ncols = range(s.ncols)

            for col in ncols:
                col_name.append(s.cell(0,col).value)

            for row in range(1,s.nrows):
                item = []
                for col in ncols:
                    item.append(s.cell(row,col).value)
                items.append(item)
        return {'items':items,'col_name':col_name}

    def file_sniff(self,target_path):
        setting = self.setting

        if not os.path.isdir(target_path):
            raise Exception("%s is not a folder!" % (target_path))

        fs = os.listdir(target_path)
        len_of_bs_type = len(setting['bitstream_filename_regex'])
        xls_file = None
        bitstream_cnt = 0
        bitsteams = []
        ignored_files = []

        # print("files in this path:",fs)
        for f in fs:
            f_detail = os.path.splitext(f)
            f_ext = f_detail[1].lower().replace('.','')
            f_name = f_detail[0]
            
            if f_ext in setting['xls_ext']:
                xls_file = f
            elif f_ext in setting['bitstream_ext']:
                if not bitstream_cnt:
                    bitstream_filename_regex_index = 0
                    bs_reg = False
                    reg_match = False

                    while (not reg_match) and (bitstream_filename_regex_index < len_of_bs_type):
                        bs_reg = re.compile(setting['bitstream_filename_regex'][bitstream_filename_regex_index][1])
                        reg_match = bs_reg.match(f_name)
                        bitstream_filename_regex_index+=1

                    bitsteams.append([int(reg_match.group(1)),f])
                    bitstream_name_format = setting['bitstream_filename_regex'][bitstream_filename_regex_index-1][0]
                elif bs_reg:
                    reg_match = bs_reg.match(f_name)
                    if not reg_match:
                        raise Exception("bitstream filename format doesn't match! Format name:%s, the file that cause this error:%s" % (bitstream_name_format,f))
                    bitsteams.append([int(reg_match.group(1)),f])
                else:
                    bitsteams.append([bitstream_cnt,f])
                bitstream_cnt+=1
            else:
                ignored_files.append(f)

        bitsteams.sort( key = lambda ele: ele[0] )
        bitsteams = [ bs[1] for bs in bitsteams ]

        if not xls_file:
            raise Exception("XLS or XLSX file not found!")

        self.xls_file = xls_file
        self.bitsteams = bitsteams
        self.bitstream_name_format = bitstream_name_format
        self.ignored_files = ignored_files

        return {
            'xls_file':xls_file,
            'bitsteams':bitsteams,
            'bitstream_name_format': bitstream_name_format,
            'ignored_files':ignored_files
        }

    def oldMeta2New(self,oldMeta):
        setting = self.setting

        col_name = []
        multiField_k = []
        multiField_match = [ __class__.remove_ptn(field,setting['xls_col2csv_col_preReplace']) for field in setting['multiField']]
        ignore_k = []
        ignore_match = [ __class__.remove_ptn(field,setting['xls_col2csv_col_preReplace']) for field in setting['ignored_col']]
        k = 0

        for col in oldMeta['col_name']:
            col_name_tmp = col.replace(setting['xls_col2csv_col_preReplace'][0],setting['xls_col2csv_col_preReplace'][1])

            for xls_col2csv_col in setting['xls_col2csv_col']:
                col_reg_match = re.match(xls_col2csv_col[0],col_name_tmp)
                if col_reg_match:
                    col_match_name = __class__.remove_ptn(col_reg_match.group(1),setting['xls_col2csv_col_preReplace'])
                    if col_match_name in ignore_match:
                        ignore_k.append(k)
                        break
                    col_name_tmp = re.sub(xls_col2csv_col[0],xls_col2csv_col[1],col_name_tmp)
                    col_name.append(col_name_tmp)
                    if col_match_name in multiField_match:
                       multiField_k.append(k)
                    break
            k+=1

        def oldDelimiter2New(cell):
            return setting["oldDelimiter2New"][1].join(
                [__class__.remove_ptn(v,setting['pattern_to_remove']).strip() for v in cell.split(setting["oldDelimiter2New"][0])]
            )

        def cell_cleanup(cell):
            return __class__.remove_ptn(cell,setting['pattern_to_remove']).strip()

        items = []

        for item in oldMeta['items']:
            items.append(
                [oldDelimiter2New(item[i]) if i in multiField_k else cell_cleanup(item[i]) for i in range(len(item)) if i not in ignore_k]
            )

        newMeta = {'items':items,'col_name':col_name}

        self.Meta = newMeta

        return newMeta

    def addBs2Meta(self,meta,bitstream):
        setting = self.setting

        col_name = meta['col_name']
        items = meta['items']
        bitstream_name_col = setting['bitstream_name_col']

        if bitstream_name_col in col_name:
            col_k = [ i for i in range(len(col_name)) if bitstream_name_col == col_name[i] ][0]
        else:
            col_k = len(col_name)

        col_name.insert(col_k,bitstream_name_col)

        if len(bitstream) is not len(items):
            raise Exception("Number of items in metadata (%d) and bitstreams (%d) doesn't match!" % (len(items),len(bitstream)))

        for i in range(len(items)):
            items[i].insert(col_k,bitstream[i])

        return {'items':items,'col_name':col_name}

    def writeCsv(self,meta,csv_file_name):
        with open(csv_file_name, 'w', encoding='utf8') as f:
            writer = csv.writer(f)
            writer.writerow(meta['col_name'])
            for item in meta['items']:
                writer.writerow(item)

    def convert(self,source_path,des_csv = False):

        try:
            print("Start to process...")
            self.file_sniff(source_path)
            self.Meta = __class__.parseXls(os.path.join(source_path,self.xls_file))
            self.oldMeta2New(self.Meta)
            self.addBs2Meta(self.Meta,self.bitsteams)
            if not des_csv:
                des_csv = os.path.join(source_path,'metadata.csv')

            self.writeCsv(self.Meta,des_csv)
            print("Result has been written to:",des_csv)
            return des_csv

        except Exception as e:
            print("\nERROR:")
            print("\t",e)
            return e
        else:
            print("Process Completed!")
        finally:
            pass

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:\n\tpython3 "+sys.argv[0]+" <path_contains_xls> [<csv_path>]")
        sys.exit(1)

    main = OldFormatToCsv()

    if len(sys.argv) >= 3:
        main.convert(sys.argv[1],sys.argv[2])
    else:
        main.convert(sys.argv[1])
