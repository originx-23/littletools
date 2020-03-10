import os
import xlrd
import xlsxwriter


def main():
    source_xls = []
    for root, dirs, files in os.walk(".", topdown=False):
        for name in files:
            target = os.path.join(root, name)
            if target.endswith(".xlsx") or target.endswith(".xls"):
                source_xls.append(target)
                # print(source_xls)
        # for name in dirs:
        #     print(os.path.join(root, name))
    # print(source_xls)
    path = os.getcwd()
    target_xls = os.path.join(path, "最爱邬仁超.xlsx")
    # 读取数据
    data = []
    for j, i in enumerate(source_xls):
        # print(j)
        wb = xlrd.open_workbook(i)
        for sheet in wb.sheets():
            # print(sheet)
            for rownum in range(sheet.nrows):
                if j > 0 and rownum == 0:
                    continue
                data.append(sheet.row_values(rownum))
    # print(data)
    # 写入数据
    workbook = xlsxwriter.Workbook(target_xls)
    worksheet = workbook.add_worksheet()
    font = workbook.add_format({"font_size": 14})
    for i in range(len(data)):
        for j in range(len(data[i])):
            worksheet.write(i, j, data[i][j], font)
    # 关闭文件流
    workbook.close()
    print("""罗玲敏爱邬仁超
名第仙人不可招
七十六年无一事
云岩深处独逍遥""")
    aa = input("朗读两遍并背诵, 按回车退出........")


if __name__ == '__main__':
    main()
