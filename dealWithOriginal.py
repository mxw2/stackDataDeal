import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import column_model
from column_model import CalculateType, ds_income_row, ds_business_cost_row, ds_business_tax_row,\
    ds_selling_expenses_row, ds_manage_expenses_row, ds_develop_expenses_row

# prepare data
book = openpyxl.load_workbook('original.xlsx')
# 数据源sheet
ds_sheet = book["利润表,资产负债表,现金流量表1"]
# 数据源开始年的start char
ds_start_year_index_char = 'B'
# 数据源开始年的end char
ds_end_year_index_char = 'J'


def create_result_sheet():
    # 判断是否存在，没有的话立刻创建，用于保存结果
    result_sheet_name = '结果'
    if result_sheet_name in book.sheetnames:
        # remove表格
        result_sheet = book[result_sheet_name]
        book.remove(result_sheet)

    # 创建一张新的sheet, 用于存放处理完毕的数据
    return book.create_sheet(result_sheet_name, 1)


def read_data():
    # 查看下有数据单元格范围，表格不是特别准确，可能是我之前wb做了很多无用功导致的，以后在试试
    # print(data_source_sheet.dimensions)
    # print("Minimum row: {0}".format(data_source_sheet.min_row))
    # print("Maximum row: {0}".format(data_source_sheet.max_row))
    # print("Minimum column: {0}".format(data_source_sheet.min_column))
    # print("Maximum column: {0}".format(data_source_sheet.max_column))

    # 读取数据源表，并且做若干判断，保证数据位置都是正确的

    # 创建结果sheet
    result_sheet = create_result_sheet()

    # result 会有（data_source_years + 1）列
    data_source_years = ord(ds_end_year_index_char) - ord(ds_start_year_index_char) + 1
    print('共有' + str(data_source_years) + '年')

    # 获取要处理的数据
    column_models = column_model.want_to_deal_with_data_source()

    # 1.从A1开始，先设置result的title row，不要和数据for-for混用，数值差1次
    for column in range(len(column_models)):
        model = column_models[column]
        result_cell_index = chr(column + ord('A')) + '1'
        result_sheet[result_cell_index] = model.name

    # 除以1亿
    yi_unit = 100000000.0

    # 开始填充result sheet，先遍历列，在遍历行
    # 2.从A2开始填充 column是0开始, A\B\C\D\E\F
    for column in range(len(column_models)):
        model = column_models[column]
        # row是从零开始的，但是在sheet是从第二行开始
        for row in range(data_source_years):
            # 有些cell不需要填充
            if data_source_years - row > model.useful_years:
                continue
            # result每次从ds中获取数据的时候，result每次row + 1，datasource每次column + 1
            ds_cell_index = chr(row + ord(ds_start_year_index_char)) + str(model.ds_row_index)
            result_cell_index = chr(column + ord('A')) + str(row + 2)
            # 逻辑判断 & 格式化数据
            if model.calculate_type == CalculateType.Year:
                year_str = ds_sheet[ds_cell_index]
                content = year_str.value.year
            elif model.calculate_type == CalculateType.OriginalData:
                original_data = ds_sheet[ds_cell_index].value
                content = two_formate(original_data / yi_unit)
            elif model.calculate_type == CalculateType.DivisionIncome:
                original_data = ds_sheet[ds_cell_index].value
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + str(ds_income_row)
                income_data = ds_sheet[income_cell_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(original_data * 100 / income_data)
            elif model.calculate_type == CalculateType.BusinessCreateProfit:
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + str(ds_income_row)
                income_data = ds_sheet[income_cell_index].value
                remind_data = ds_sheet[income_cell_index].value
                business_cost_index = chr(row + ord(ds_start_year_index_char)) + str(ds_business_cost_row)
                remind_data -= ds_sheet[business_cost_index].value
                ds_business_tax_index = chr(row + ord(ds_start_year_index_char)) + str(ds_business_tax_row)
                remind_data -= ds_sheet[ds_business_tax_index].value
                ds_selling_expenses_index = chr(row + ord(ds_start_year_index_char)) + str(ds_selling_expenses_row)
                remind_data -= ds_sheet[ds_selling_expenses_index].value
                ds_manage_expenses_index = chr(row + ord(ds_start_year_index_char)) + str(ds_manage_expenses_row)
                remind_data -= ds_sheet[ds_manage_expenses_index].value
                ds_develop_expenses_index = chr(row + ord(ds_start_year_index_char)) + str(ds_develop_expenses_row)
                remind_data -= ds_sheet[ds_develop_expenses_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(remind_data * 100 / income_data)
            elif model.calculate_type == CalculateType.GrossMargin:
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + str(ds_income_row)
                income_data = ds_sheet[income_cell_index].value
                remind_data = ds_sheet[income_cell_index].value
                business_cost_index = chr(row + ord(ds_start_year_index_char)) + str(ds_business_cost_row)
                remind_data -= ds_sheet[business_cost_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(remind_data * 100 / income_data)
            else:
                content = ''

            result_sheet[result_cell_index] = content
            cell = result_sheet[result_cell_index]
            # 统一设置文字颜色等
            if model.text_color is not None:
                cell.fill = PatternFill("solid", fgColor=model.text_color)

    # 保存下
    book.save('original.xlsx')


# 返回两位小数数字
def two_formate(income):
    result_str = str(income).split('.')[0] + '.' + str(income).split('.')[1][:2]
    return float(result_str)


if __name__ == '__main__':
    read_data()