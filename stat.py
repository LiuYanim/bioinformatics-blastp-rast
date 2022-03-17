import xlrd
import xlwt

if __name__ == "__main__":
    data = xlrd.open_workbook('/Users/Lenovo/Desktop/whnfull并集无重复 - 副本.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    tmp = dict()
    for i in range(nrows):
        if i == 0:
            continue
        row = table.row_values(i)
        key = row[0]
        if key in tmp:
            tmp[key].append(row[1])
        else:
            tmp[key] = [row[1]]

    workbook = xlwt.Workbook()
    result_sheet = workbook.add_sheet('result', cell_overwrite_ok=True)
    tem = 0
    for key in tmp.keys():
        result_sheet.write(tem, 0, key)
        s = ""
        for i, content in enumerate(tmp[key]):
            s = s + content + ", "
        result_sheet.write(tem, 1, s)
        result_sheet.write(tem, 2, len(tmp[key]))
        tem = tem + 1

    workbook.save('/Users/Lenovo/Desktop/1stat-whnfull.xlsx')