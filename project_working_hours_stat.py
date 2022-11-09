import os
import re

import xlrd
import xlwt


class stat_data:
    name = ''
    total_working_month = 0.0
    comments = ''


def read_project_folder(folder, result):
    for _, _, files in os.walk(folder):
        for file in files:
            # the file does NOT start with ~$, and end with .xlsx
            if str.find(file, '~$') < 0 and str.find(file, '.xls') > 0:
                read_each_project_measure_data(folder + "/" + file, result)


def read_each_project_measure_data(file, result):
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_name('sheet1')
    total = sheet.nrows
    project_name = re.search("(?<=/)\\w*(?=_)", file).group(0)
    for x in range(1, total - 1):
        name = sheet.cell_value(x, 2)
        working_month = sheet.cell_value(x, 4)
        if working_month == 0:
            continue
        data = result.get(name, stat_data())
        data.name = name
        data.total_working_month += float(working_month)
        data.comments += project_name + '(' + str(working_month) + ');'
        result[name] = data

    pass


if __name__ == '__main__':
    print("start")
    project_measure_folder = 'C:/Users/user/Desktop/3'
    result = {}
    read_project_folder(project_measure_folder, result)

    # for key in result.keys():
    #     print(result[key].name + " " + str(result[key].total_working_month) + " " + result[key].comments)

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("sheet1")

    sheet.write(0, 0, '姓名')
    sheet.write(0, 1, '工作量(人月)')
    sheet.write(0, 2, '备注')
    n = 1
    for key in result.keys():
        data = result[key]
        sheet.write(n, 0, data.name)
        sheet.write(n, 1, data.total_working_month)
        sheet.write(n, 2, data.comments)
        n += 1

    workbook.save(project_measure_folder + '/result.xls')
