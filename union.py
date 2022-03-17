import xlrd
import xlwt

father = list()
people = list()
rank = list()


def find(x):
    if x != father[x]:
        father[x] = find(father[x])
    return father[x]


def union(x, y):
    x = find(x)
    y = find(y)

    if rank[x] > rank[y]:
        father[y] = x
        if x != y:
            people[x].extend(people[y])
    else:
        father[x] = y
        if x != y:
            people[y].extend(people[x])

        if rank[x] == rank[y]:
            rank[y] += 1


if __name__ == "__main__":
    data = xlrd.open_workbook('/Users/Lenovo/Desktop/whndb（无重复）.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    total = set()
    pair = list()
    for i in range(nrows):
        if i == 0:
            continue
        row = table.row_values(i)
        total.add(row[0])
        total.add(row[1])

    father_tmp = list(total)
    father = [i for i, content in enumerate(father_tmp)]

    for i in range(nrows):
        if i == 0:
            continue
        row = table.row_values(i)
        if row[0] != row[1]:
            pair.append([father_tmp.index(row[0]), father_tmp.index(row[1])])

    for content in father:
        people.append([content])
        rank.append(0)

    for pair_content in pair:
        union(pair_content[0], pair_content[1])

    workbook = xlwt.Workbook()
    result_sheet = workbook.add_sheet('result', cell_overwrite_ok=True)
    tem = 0
    for i, content in enumerate(father):
        if find(i) == i:
            for j, item in enumerate(people[i]):
                if father_tmp[content] != father_tmp[item]:
                    result_sheet.write(tem, 0, father_tmp[content])
                    result_sheet.write(tem, 1, father_tmp[item])
                    tem += 1

    workbook.save('/Users/Lenovo/Desktop/1whndb（无重复）.xls')

