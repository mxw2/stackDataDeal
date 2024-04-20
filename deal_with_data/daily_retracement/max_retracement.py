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

    print("Maximum row: {0}, 有效行数量:{1}".format(ds_sheet.max_row, ds_sheet.max_row - 1))
    print("Maximum column: {0}".format(ds_sheet.max_column))

    price_infos = []
    # 遍历每一行数据
    # range(5):  # 从0到4
    # range(1, 6):  # 从1开始，到5结束（5会被包含在内）
    for i in range(2, ds_sheet.max_row + 1):
        min_value = ds_sheet[column_index_for_min_value + str(i)].value
        max_value = ds_sheet[column_index_for_max_value + str(i)].value
        price_info = DailyPriceInfo(max_value, min_value)
        price_infos.append(price_info)
        print("D{0}, 最大值:{1}, 最小值:{2}, 平均值:{3}".format(str(i),
                                                                max_value,
                                                                min_value,
                                                                price_info.average_price))

    # 按照时间由远到近的排序
    price_infos.reverse()

    # 统计最大亏损
    max_loss = 0
    max_loss_percent = 0
    # 遍历每个价格 & 求出连续5天最大回撤
    for i in range(len(price_infos)):
        current_price = price_infos[i].average_price
        for j in range(i + 1, i + target_days):
            # 特殊处理，保证j < price_infos.count
            if j >= len(price_infos):
                print("遍历价格数组越界打印：i :{0}, j :{1}，continue".format(i, j))
                continue
            other_price = price_infos[j].average_price
            if other_price < current_price:
                temp_loss = other_price - current_price
                max_loss = min(temp_loss, max_loss)
                max_loss_percent = min(temp_loss / current_price, max_loss_percent)

    # 打印最后的结果哈
    print("最大损失: {0}, 最大回撤: {1}".format(max_loss, max_loss_percent))

# ********************************** 启动 **********************************
if __name__ == '__main__':
    read_data()