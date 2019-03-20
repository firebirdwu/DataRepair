# author:wufei
import xlrd, xlwt
import os
import sys


def scan_excel(excel_path):
    if not os.path.exists(excel_path):
        print("文件未找到" + excel_path)
    work_book = xlrd.open_workbook(excel_path)
    # print(work_book.sheet_names())
    data_sheet = work_book.sheet_by_index(0)
    # print(data_sheet.name)
    rowNum = data_sheet.nrows
    colNum = data_sheet.ncols
    r_row = 1

    while r_row < rowNum:
        rows = data_sheet.row_values(r_row)
        r_col = 0
        while r_col < colNum:
            print(rows[r_col])
            r_col = r_col + 1
        r_row = r_row + 1


if __name__ == '__main__':
    project_path = os.path.dirname(sys.path[0])
    excel_path = os.path.join(project_path, 'excel/excel_test.xlsx')
    scan_excel(excel_path)
