import openpyxl
from openpyxl import Workbook


def read_data():
    # 读取数据
    book = openpyxl.load_workbook('original.xlsx')

    data_source_sheet = book["利润表,资产负债表,现金流量表1"]
    print(data_source_sheet.title)

    # 查看下有数据单元格范围，表格不是特别准确，可能是我之前wb做了很多无用功导致的，以后在试试
    print(data_source_sheet.dimensions)
    print("Minimum row: {0}".format(data_source_sheet.min_row))
    print("Maximum row: {0}".format(data_source_sheet.max_row))
    print("Minimum column: {0}".format(data_source_sheet.min_column))
    print("Maximum column: {0}".format(data_source_sheet.max_column))

    # 判断是否存在，没有的话立刻创建，用于保存结果
    result_sheet_name = '结果'
    #
    if result_sheet_name in book.sheetnames:
        # remove表格
        result_sheet = book[result_sheet_name]
        book.remove(result_sheet)

    # 创建一张新的sheet, 用于存放处理完毕的数据
    result_sheet = book.create_sheet(result_sheet_name, 1)
    # book.save('original.xlsx')

    # 做一些准备工作
    start_year_index_str = 'B'
    # 未来自动寻找结束年
    end_year_index_str = 'J'
    data_source_years = ord(end_year_index_str) - ord(start_year_index_str) + 1
    print('共有' + str(data_source_years) + '年')

    # 获取开始年份
    start_year_str = data_source_sheet[start_year_index_str + '5']
    print('start_year_str = ' + str(start_year_str.value.year))

    result_sheet['A1'] = '年份'
    # title_names = ['年份', '营业收入']
    for index in range(data_source_years):
        current_row = 'A' + str(index + 2)
        result_sheet[current_row] = str(start_year_str.value.year + index)

    # 保存下
    book.save('original.xlsx')

    # 具体科目
    # 营业收入


if __name__ == '__main__':
    read_data()