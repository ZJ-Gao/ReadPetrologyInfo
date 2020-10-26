'''
    Author: ZIJIE GAO
    Date: 4/19/2019
    Function: Read the well-name, depth, sequence name, and content of quartz, feldspar, rock fragments
    chlorite and calcite from Excel documents.
    Version: 3.0
    Update the codes with module can read multiple Excel documents
'''

#encoding=utf-8

import xlrd
import os

def get_filename(path,filetype):
    '''
        Read the file name
    '''
    name = []
    for root,dirs,files in os.walk(path):
        for i in files:
            if filetype in i:
                name.append(i.replace(filetype,''))
    return name

def write_txt(list):
    '''
        Write data into a new .txt document
    '''
    fileObject = open('thin_section.txt', 'a', encoding='utf-8')
    fileObject.write('\n')
    for ip in list:
        fileObject.writelines(str(ip))
        fileObject.write(' ')
    fileObject.write('\n')

    fileObject.close()

def main():

    list_final = []
    path = 'C:\\changename_two'
    filetype = '.xlsx'
    name = get_filename(path, filetype)
    for word in name:
        data = xlrd.open_workbook(word + '.xlsx')
        table = data.sheets()[0]
        col_0 = table.col_values(0)
        for i, element in enumerate(col_0):
            if element == '砂 岩 薄 片 鉴 定 报 告':
                well_name = table.cell(i + 2, 2).value
                strata = table.cell(i + 1, 7).value
                depth = table.cell(i + 2, 7).value
                quartz = table.cell(i + 7, 0).value
                chert = table.cell(i + 7, 1).value
                if chert == '':
                    chert = 0
                feldspar = table.cell(i + 7, 2).value
                # Content of rock fragments
                row7 = table.row_values(i + 7)
                # Pick out Null value
                while '' in row7:
                    row7.remove('')
                # print(row7)
                sum_rol_content = round(sum(row7),1)
                list_final = [word, well_name, strata, depth, quartz, chert, feldspar, sum_rol_content]
                print(list_final)
                write_txt(list_final)

if __name__ == '__main__':
        main()
