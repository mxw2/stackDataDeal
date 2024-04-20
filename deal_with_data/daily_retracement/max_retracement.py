print('开始统计当前股票的最大回撤')

from daily_price_info import DailyPriceInfo
import openpyxl

# ********************************** 常量 **********************************

# 一周5个交易日
target_days = 5

workbook_name = '108sm-ecmpp.xlsx'
# prepare data
book = openpyxl.load_workbook(workbook_name)
# 数据源sheet
ds_sheet = book["108sm-ecmpp"]

# 从第二行开始读数据
row_index_for_start_read = 2
column_index_for_max_value = 'D'
column_index_for_min_value = 'E'

# ********************************** 读取表内容 **********************************
def read_data():
    # 读取数据源表，并且做若干判断，保证数据位置都是正确的
    print("Maximum row: {0}".format(ds_sheet.max_row))
    print("Maximum column: {0}".format(ds_sheet.max_column))

    price_infos = []
    # 遍历每一行数据, range(5):  # 从0到4
    # row是从1开始的，所以遍历的时候，必须要遍历到max_row + 1
    # 例如：共11行数据，要遍历第2行-第11行数据
    for i in range(2, ds_sheet.max_row + 1):
        min_value = ds_sheet[column_index_for_min_value + str(i)].value
        max_value = ds_sheet[column_index_for_max_value + str(i)].value
        price_info = DailyPriceInfo(max_value, min_value)
        price_infos.append(price_info)
        print("D{0}, 最大值:{1}, 最小值:{2}, 平均值:{3}".format(str(i),
                                                                max_value,
                                                                min_value,
                                                                price_info.average_price))

# ********************************** 启动 **********************************
if __name__ == '__main__':
    read_data()