#-*- coding:utf-8 -*-

import os
import sys
import xlrd

reload(sys)
sys.setdefaultencoding('utf-8')

def Usage():
    print '''
    ###########################################################
    python excel_txt.py filename
    ###########################################################
    '''
    sys.exit(0)


def xlsx_to_txt(fname):
    '''
    translate excel file to txt file.
    '''
    # get input filename
    fwname = os.path.basename(os.path.splitext(fname)[0])

    # open excel and read row data.
    data = xlrd.open_workbook(fname)
    sheetnames = data.sheet_names()
    for m in sheetnames:
        table = data.sheet_by_name(m)
        nrows = table.nrows
        ncols = table.ncols
        for i in range(0,nrows):
            data_row = table.row_values(i)
            data_row_list = []
            for j in data_row:
                # cov different encoding to unicode.it can extends other encoding.
                if isinstance(j,float):
                    j = str(j).decode('utf-8')
                data_row_list.append(j)

            #data_row = map(str,data_row)
            str_data_row = ','.join(data_row_list)
            with open(fwname + '.txt',"a+") as fw:
                fw.write(str_data_row + '\n')

if __name__ == "__main__":
    if len(sys.argv) != 2:
        Usage()
    fname = sys.argv[1]
    xlsx_to_txt(fname)
