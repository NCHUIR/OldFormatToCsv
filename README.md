OldFormatToCsv
==============

將舊格式轉成 Dspace4 匯出的 metadata 格式 (包含 bitstream 格式)，可再透過 csvToDspaceSaf 轉成可匯入之格式

## Requirement

 * Python3
 * xlrd module,use `pip install xlrd` to get xlrd

## Usage

```
python3 OldFormatToCsv.py <path_contains_xls> [<csv_path>]
``` 

## About OldFormat (NTUR format) to CSV format

 * xls / xlsx metadata file => output csv file
 * column name: `column_name=language` => `dc.column_name[language]`
    * `:` will be changed to `.` in column name
 * these columns will be ignored:`sys_filename`,`sys_replace`
 * multi value in one cell delimitor: `;` => `||`
 
## contents sniffer

 * use the following file name formats to detect the the_order in metadata csv file
     - `handle_or_anything-the_order.pdf`
     - `the_order.pdf`
 * store file name in csv's `contents` column

### About Author

 * PastLeo
 * 2014/8/7
