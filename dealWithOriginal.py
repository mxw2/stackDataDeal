import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment


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
    if result_sheet_name in book.sheetnames:
        # remove表格
        result_sheet = book[result_sheet_name]
        book.remove(result_sheet)

    # 创建一张新的sheet, 用于存放处理完毕的数据
    result_sheet = book.create_sheet(result_sheet_name, 1)

    # 做一些准备工作
    start_year_index_str = 'B'
    # 未来自动寻找结束年
    end_year_index_str = 'J'
    data_source_years = ord(end_year_index_str) - ord(start_year_index_str) + 1
    print('共有' + str(data_source_years) + '年')

    # 获取开始年份
    start_year_str = data_source_sheet[start_year_index_str + '5']
    print('start_year_str = ' + str(start_year_str.value.year))

    ds_business_income_cell_row = '12'  # 营业收入
    ds_cash_flow_cell_row = '244'  # '经营活动现金流入'
    ds_net_profit_cell_row = '62'  # '净利润'
    ds_net_cash_cell_row = '275'  # '经营活动产生的现金流量净额'
    ds_net_money_in_cell_row = '318'  # '筹资净额（亿元）'
    ds_net_money_out_cell_row = '297'  # '投资净额（亿元）'
    ds_operating_cost_cell_row = '19'  # '营业成本%'

    # 配置第一行的titles
    first_row = ['年份',
                 '营业收入（亿元）',
                 '经营活动现金流入（亿元）',
                 '净利润（亿元）',
                 '经营活动产生的现金流量净额（亿元）',
                 '筹资净额（亿元）',
                 '投资净额（亿元）',
                 '营业成本%',
                 '营业税金及附加%']
    result_sheet.append(first_row)

    yi_unit = 100000000.0
    # 建议修改，数组都以1开始即可
    # 配置每一列数据
    for index in range(data_source_years):
        # 数据源中开始遍历的列的index值
        ds_begin_colume_char = 'B'
        # 1列：拼写当前年
        result_sheet['A' + str(index + 2)] = str(start_year_str.value.year + index)

        ds_current_column_char = ord(ds_begin_colume_char) + index
        # 2列：填写营业收入
        income_true = data_source_sheet[chr(ds_current_column_char) + ds_business_income_cell_row].value / yi_unit
        result_sheet['B' + str(index + 2)] = two_formate(income_true)

        # 3列：营业收入
        income = data_source_sheet[chr(ds_current_column_char) + ds_cash_flow_cell_row].value / yi_unit
        result_sheet['C' + str(index + 2)] = two_formate(income)

        # 4列:净利润
        income = data_source_sheet[chr(ds_current_column_char) + ds_net_profit_cell_row].value / yi_unit
        result_sheet['D' + str(index + 2)] = two_formate(income)

        # 5列:经营活动产生的现金流量净额【要10年 & 要6年】
        income = data_source_sheet[chr(ds_current_column_char) + ds_net_cash_cell_row].value / yi_unit
        result_sheet['E' + str(index + 2)] = two_formate(income)

        if data_source_years - index <= 6:
            cell = result_sheet['E' + str(index + 2)]
            cell.fill = PatternFill("solid", fgColor="ff0000")
            # 6列:筹资净额【只要6年的】
            income = data_source_sheet[chr(ds_current_column_char) + ds_net_money_in_cell_row].value / yi_unit
            result_sheet['F' + str(index + 2)] = two_formate(income)
            cell = result_sheet['F' + str(index + 2)]
            cell.fill = PatternFill("solid", fgColor="ff0000")

            # 7列:投资净额【只要6年的】
            income = data_source_sheet[chr(ds_current_column_char) + ds_net_money_out_cell_row].value / yi_unit
            result_sheet['G' + str(index + 2)] = two_formate(income)
            cell = result_sheet['G' + str(index + 2)]
            cell.fill = PatternFill("solid", fgColor="ff0000")

        if data_source_years - index <= 4:
            # 8列:'营业成本%'【只要4年】
            income = data_source_sheet[chr(ds_current_column_char) + ds_operating_cost_cell_row].value / yi_unit
            result_sheet['H' + str(index + 2)] = two_formate(income / income_true * 100)

    # 保存下
    book.save('original.xlsx')

# 目前存储的是字符串，有待修改
def two_formate(income):
    return str(income).split('.')[0] + '.' + str(income).split('.')[1][:2]


if __name__ == '__main__':
    read_data()