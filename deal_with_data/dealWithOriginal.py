import openpyxl
from openpyxl.styles import PatternFill
import column_model
from column_model import CalculateType, ds_income_string, ds_business_cost_string, ds_business_tax_string, \
    ds_selling_expenses_string, ds_manage_expenses_string, ds_develop_expenses_string, ds_row_for_key

# 使用步骤：
# 1.将要处理的excle放到python脚本同级目录下
# 2.修改python脚本中workbook_name为你的文件名称
# 3.确定ds_sheet表名是否相同
# 4.确定最后算出来的"共有x年"与你的一致，因为如果表格是你手动修改过，添加过超过原来最大列的cell值，算出来的就是有问题的
workbook_name = '黄山旅游08-18.xlsx'
# prepare data
book = openpyxl.load_workbook(workbook_name)
# 数据源sheet
ds_sheet = book["利润表,资产负债表,现金流量表3"]
# 数据源开始年的start char
ds_start_year_index_char = 'B'
# 数据源中 key:value
ds_first_column_dictionary = {}


def create_result_sheet():
    # 判断是否存在，没有的话立刻创建，用于保存结果
    result_sheet_name = '结果'
    if result_sheet_name in book.sheetnames:
        # remove表格
        result_sheet = book[result_sheet_name]
        book.remove(result_sheet)
        print('remove掉已有的result sheet')
    else:
        print('没有[结果]sheet，需要手动新建')
    # 创建一张新的sheet, 用于存放处理完毕的数据
    return book.create_sheet(result_sheet_name, 1)


def create_first_column_dictionary():
    print("Maximum row: {0}".format(ds_sheet.max_row))
    # key:
    for row in range(ds_sheet.max_row):
        # print('current row ' + str(row + 1))
        current_row = row + 1;
        cell = ds_sheet["A" + str(current_row)]
        if cell.value is None:
            continue
        elif len(cell.value) > 0:
            value = cell.value.strip()
            ds_first_column_dictionary[value] = current_row
            print('key:' + value + ', value:' + str(current_row))
        else:
            continue


def suitable_result_column(column):
    has_extra_a = column / 26
    # print('has_extra_a = ' + str(has_extra_a))
    # 1/26= 0.0x
    if has_extra_a < 1:
        suitable_column_str = chr(column + ord('A'))
    else:
        deal_with_column = column % 26
        # 目前只处理A1 & AA1情况，BA1+不再考虑
        suitable_column_str = 'A' + chr(deal_with_column + ord('A'))
    # print('result.column = ' + str(column) + ',suitable.column = ' + suitable_column_str)
    return suitable_column_str


def read_data():
    # 读取数据源表，并且做若干判断，保证数据位置都是正确的
    print("Maximum column: {0}".format(ds_sheet.max_column))

    # 维护一个字典，保存"寻找的字符串":"row"
    create_first_column_dictionary()

    # 创建结果sheet
    result_sheet = create_result_sheet()
    print('result sheet获取成功 ', result_sheet)

    # result 会有（data_source_years + 1）列
    # 去除第一个A列
    data_source_years = ds_sheet.max_column - 1
    print('共有' + str(data_source_years) + '年')

    # 获取要处理的数据
    column_models = column_model.want_to_deal_with_data_source()

    # 1.从A1开始，先设置result的title row，不要和数据for-for混用，数值差1次
    print('先将title安置到第一行')
    # 超过26列，就有问题。[1了。。。
    for column in range(len(column_models)):
        model = column_models[column]
        result_cell_index = suitable_result_column(column) + '1'
        print('title = ' + model.name + 'result_cell_index = ' + str(result_cell_index))
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
            actually_userful_years = model.useful_years
            if model.useful_years == column_model.useful_years_default:
                actually_userful_years = data_source_years

            if data_source_years - row > actually_userful_years:
                continue
            # result每次从ds中获取数据的时候，result每次row + 1，datasource每次column + 1
            ds_cell_index = chr(row + ord(ds_start_year_index_char)) + str(model.ds_row_index)
            result_cell_index = suitable_result_column(column) + str(row + 2)
            print('双层for循环，转换index。ds_index[' + ds_cell_index + ':' + result_cell_index + ']result_index')
            # 逻辑判断 & 格式化数据
            if model.calculate_type == CalculateType.Year:
                year_str = ds_sheet[ds_cell_index]
                content = year_str.value.year
                print('进入到CalculateType.Year:' + str(content))
            elif model.calculate_type == CalculateType.OriginalData:
                original_data = ds_sheet[ds_cell_index].value
                content = two_formate(original_data / yi_unit)
            elif model.calculate_type == CalculateType.DivisionIncome:
                original_data = ds_sheet[ds_cell_index].value
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_income_string))
                income_data = ds_sheet[income_cell_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(original_data * 100 / income_data)
            elif model.calculate_type == CalculateType.BusinessCreateProfit:
                # 建议采用数组更好管理
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_income_string))
                income_data = ds_sheet[income_cell_index].value
                remind_data = ds_sheet[income_cell_index].value
                business_cost_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_business_cost_string))
                remind_data -= ds_sheet[business_cost_index].value
                ds_business_tax_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_business_tax_string))
                remind_data -= ds_sheet[ds_business_tax_index].value
                ds_selling_expenses_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_selling_expenses_string))
                remind_data -= ds_sheet[ds_selling_expenses_index].value
                ds_manage_expenses_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_manage_expenses_string))
                remind_data -= ds_sheet[ds_manage_expenses_index].value
                ds_develop_expenses_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_develop_expenses_string))
                remind_data -= ds_sheet[ds_develop_expenses_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(remind_data * 100 / income_data)
            elif model.calculate_type == CalculateType.GrossMargin:
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_income_string))
                income_data = ds_sheet[income_cell_index].value
                remind_data = ds_sheet[income_cell_index].value
                business_cost_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_business_cost_string))
                remind_data -= ds_sheet[business_cost_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(remind_data * 100 / income_data)
            elif model.calculate_type == CalculateType.CommonTurnoverRate or model.calculate_type == CalculateType.GoodsTurnoverRate:
                # B12 / (B153 + A153) / 2
                # B19 / (B153 + A153) / 2
                if model.calculate_type == CalculateType.CommonTurnoverRate:
                    ds_target_string = ds_income_string
                else:
                    ds_target_string = ds_business_cost_string
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + str(ds_row_for_key(ds_target_string))
                income_data = ds_sheet[income_cell_index].value
                # B153 & A153
                common_data = ds_sheet[ds_cell_index].value
                left_common_cell_index = chr(row + ord(ds_start_year_index_char) - 1) + str(model.ds_row_index)
                left_common_data = ds_sheet[left_common_cell_index].value
                content = two_formate(income_data / ((left_common_data + common_data) / 2.0))
            else:
                content = ''

            result_sheet[result_cell_index] = content
            cell = result_sheet[result_cell_index]
            # 统一设置文字颜色等
            if model.text_color is not None:
                cell.fill = PatternFill("solid", fgColor=model.text_color)

    # 保存下
    book.save(workbook_name)


# 返回两位小数数字
def two_formate(income):
    result_str = str(income).split('.')[0] + '.' + str(income).split('.')[1][:2]
    return float(result_str)


if __name__ == '__main__':
    read_data()
