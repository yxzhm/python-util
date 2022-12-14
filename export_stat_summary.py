import xlrd
import xlwt


class stat_data:
    name = ''
    total_working_month = 0.0
    comments = ''


def get_data(file, result):
    book = xlrd.open_workbook(file)
    sheet = book.sheet_by_index(0)
    total = sheet.nrows

    for x in range(1, total):
        name = sheet.cell_value(x, 1)
        name = name.replace("t_", "")
        working_month = sheet.cell_value(x, 3)
        project_name = sheet.cell_value(x, 2)
        if working_month == 0:
            continue
        data = result.get(name, stat_data())
        data.name = name
        data.total_working_month += float(working_month/176)
        data.comments += project_name + '(' + str(round(working_month/176, 2)) + ');'
        result[name] = data


if __name__ == '__main__':
    print("start")
    result = {}

    get_data('C:/Users/user/Desktop/20221117）.xlsx', result)
    #get_data('C:/Users/user/Desktop/4/6-7.工作量-统计人员在项目的工时2-9.xlsx', result)

    # for key in result.keys():
    #     data = result[key]
    #     print(data.name + "\t" + str(data.total_working_month) + "\t" + data.comments)



    facebook = xlrd.open_workbook('C:/Users/user/Desktop/20221116.xlsx')
    facebook_sheet = facebook.sheet_by_index(0)
    n=0
    outputbook = xlwt.Workbook()
    output = outputbook.add_sheet("sheet1")
    for i in range(1,facebook_sheet.nrows):
        name = facebook_sheet.cell_value(i,1)
        output.write(n, 0, name)
        data = result.get(name, stat_data())
        output.write(n,1, round(data.total_working_month,2))
        output.write(n,2,data.comments)
        n+=1
    outputbook.save('C:/Users/user/Desktop/3/result.xls')
