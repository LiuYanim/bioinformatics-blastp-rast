import xlrd
import xlwt

if __name__ == "__main__":
    data = xlrd.open_workbook('/home/xiaoju/data1.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows
    tmp = dict()
    result = dict()
    for i in range(nrows):
        if i == 0:
            continue
        row = table.row_values(i)
        key = row[0] + ':' + row[1]
        if key in tmp:
            if float(row[10]) < tmp.get(key):
                tmp[key] = float(row[10])
                result[key] = i
        else:
            tmp[key] = float(row[10])
            result[key] = i

    workbook = xlwt.Workbook()
    result_sheet = workbook.add_sheet('result', cell_overwrite_ok=True)
    tem = 0
    for key in result.keys():
        t_row = table.row_values(result[key])
        for i, content in enumerate(t_row):
            result_sheet.write(tem, i, content)
        tem = tem + 1

    workbook.save('/home/xiaoju/result.xlsx')
