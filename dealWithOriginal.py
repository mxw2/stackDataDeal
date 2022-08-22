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

    # 营业收入
    ds_business_income_cell_tuple = tuple(['B', 12])


    # 配置第一行的titles
    first_row = ['年份', '营业收入', '经营活动现金流入', '净利润', '经营活动产生的现金流量净额', '筹资净额', '投资净额', '营业成本%', '营业税金及附加%']
    # for index in range(len(first_row)):
        # result_sheet['B' + '1'] = first_row[index]
    result_sheet.append(first_row)
    # 配置每一列数据
    for index in range(data_source_years):
        current_row = 'A' + str(index + 2)
        result_sheet[current_row] = str(start_year_str.value.year + index)

        # 填写营业收入
        test = ds_business_income_cell_tuple[0]
        # 数据源test需要加1
        print('B12=', data_source_sheet['B12'].value)
        # print('chr(ord(test) + 1) + 12', chr(ord(test) + index) + '12')
        income = data_source_sheet[chr(ord(test) + index) + '12'].value / 100000000.0
        format_in_come = str(income).split('.')[0] + '.' + str(income).split('.')[1][:2]
        result_sheet[test + str(index + 2)] = format_in_come

        print('chr(ord(test) + 1) + 12 = ', income)

    # 保存下
    book.save('original.xlsx')

    # 具体科目
    # 营业收入


if __name__ == '__main__':
    read_data()