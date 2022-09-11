from openpyxl.styles import PatternFill
import column_model
from column_model import *
from openpyxl.chart import *  # BarChart, Reference
from openpyxl.drawing.fill import PatternFillProperties
from openpyxl.chart.label import *

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
            if model.useful_years == column_model.useful_years_max:
                actually_userful_years = data_source_years

            if data_source_years - row > actually_userful_years:
                continue
            # result每次从ds中获取数据的时候，result每次row + 1，datasource每次column + 1
            ds_cell_index = chr(row + ord(ds_start_year_index_char)) + model.ds_row_index_string
            # ds_cell_index = ds_row_for_key(ds_cell_string)
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
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_income_string)
                income_data = ds_sheet[income_cell_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(original_data * 100 / income_data)
            elif model.calculate_type == CalculateType.BusinessCreateProfit:
                # 建议采用数组更好管理
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_income_string)
                income_data = ds_sheet[income_cell_index].value
                remind_data = ds_sheet[income_cell_index].value
                business_cost_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_business_cost_string)
                remind_data -= ds_sheet[business_cost_index].value
                ds_business_tax_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_business_tax_string)
                remind_data -= ds_sheet[ds_business_tax_index].value
                ds_selling_expenses_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_selling_expenses_string)
                remind_data -= ds_sheet[ds_selling_expenses_index].value
                ds_manage_expenses_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_manage_expenses_string)
                remind_data -= ds_sheet[ds_manage_expenses_index].value
                ds_develop_expenses_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_develop_expenses_string)
                remind_data -= ds_sheet[ds_develop_expenses_index].value
                # 先乘上100在除数，保证数据格式稍微好看点
                content = two_formate(remind_data * 100 / income_data)
            elif model.calculate_type == CalculateType.GrossMargin:
                # 获取当年的营业收入cell索引
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_income_string)
                income_data = ds_sheet[income_cell_index].value
                remind_data = ds_sheet[income_cell_index].value
                business_cost_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_business_cost_string)
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
                income_cell_index = chr(row + ord(ds_start_year_index_char)) + ds_row_for_key(ds_target_string)
                income_data = ds_sheet[income_cell_index].value
                # B153 & A153
                common_data = ds_sheet[ds_cell_index].value
                left_common_cell_index = chr(row + ord(ds_start_year_index_char) - 1) + model.ds_row_index_string
                left_common_data = ds_sheet[left_common_cell_index].value
                content = two_formate(income_data / ((left_common_data + common_data) / 2.0))
            else:
                content = ''

            result_sheet[result_cell_index] = content
            cell = result_sheet[result_cell_index]
            # 统一设置文字颜色等
            if model.text_color is not None:
                cell.fill = PatternFill("solid", fgColor=model.text_color)
    save_result_sheet()

    # 创建图标
    chart = BarChart()
    chart.dLbls = DataLabelList()
    # chart.dLbls.showCatName = True  # 标签显示(显示第几列的 1,-23  显示不符合预期)
    chart.dLbls.showVal = True  # 数量显示

    # 1.先设置Y轴，看看每年的数据
    values = Reference(result_sheet, min_col=6, min_row=7, max_col=8, max_row=14)
    # s.marker.symbol = 'circle'
    # true 表示数据中不包含横向的titles()
    chart.add_data(data=values, titles_from_data=False)

    # 2.放到后边才能x轴正常出现年的时间
    cats = Reference(result_sheet, min_col=1, min_row=7, max_col=1, max_row=14)
    chart.set_categories(cats)
    # 修改表格背景的
    # chart.plot_area.graphicalProperties = GraphicalProperties(solidFill="999999")
    s0 = chart.series[0]
    # 设置线条颜色，不设置时默认为砖红色,其他颜色代码可参见 https://www.fontke.com/tool/rgb/00aaaa/
    # s2.graphicalProperties.line.solidFill = "FF5555"
    # s2.graphicalProperties.line.width = 30000  # 控制线条粗细
    # fill0 = PatternFillProperties()
    # fill0.background = ColorChoice(prstClr="red")
    s0.graphicalProperties.solidFill = 'FF0000'

    s1 = chart.series[1]
    # fill1 = PatternFillProperties(prst='wave')
    # fill1.background = ColorChoice(prstClr="blue")
    s1.graphicalProperties.solidFill = '0938F7'

    s2 = chart.series[2]
    # fill2 = PatternFillProperties()
    # fill2.background = ColorChoice(prstClr="orange")
    s2.graphicalProperties.solidFill = 'EE9611'

    chart.width = 32
    chart.height = 16
    chart.type = 'col'
    # 自定义风格时候，不要用这个字段
    # chart.style = 15
    chart.shape = 4
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = '历年现金流量净额（单位：亿元）'
    chart.x_axis.tickLblPos = 'low'
    result_sheet.add_chart(chart, 'A27')
    save_result_sheet()


# 返回两位小数数字
def two_formate(income):
    result_str = str(income).split('.')[0] + '.' + str(income).split('.')[1][:2]
    return float(result_str)


if __name__ == '__main__':
    read_data()
