import time

import openpyxl

# ************************** 读取数据源 sheet **************************************
workbook_name = 'ZKE30N-202301'
workbook_name_xlsx = workbook_name + '.xlsx'
# prepare data
book = openpyxl.load_workbook(workbook_name_xlsx, data_only=True)
# 数据源sheet
ds_sheet_name = "信创表"
ds_sheet = book[ds_sheet_name]
# 数据源sheet中具体列的名称,需要对他做过滤
ds_sheet_target_column_name = '产品归属'
# 产品归属这一列中做过滤的可能数组
ds_sheet_target_items = ['X86企业级', 'X86终端', '0']
print("ds sheet max column: {0}".format(ds_sheet.max_column))
print("ds sheet max row: {0}".format(ds_sheet.max_row))

# ************************** 创建 filter target sheet **************************************

result_work_book = openpyxl.Workbook()
result_work_sheet = result_work_book.active

# 将数据源第一行数据copy到结果表格的第一行
# https://fishc.com.cn/thread-178713-1-1.html
# openpyxls 选择整行后会变成一个cell组成的tuple对象，获取数据需要逐个获取：
ds_first_row_data = [item.value for item in ds_sheet[1]]
result_work_sheet.append(ds_first_row_data)

# 从数据源第二行开始读取，看看哪个属于"ds_sheet_target_items"
# "产品归属"这一列在G1
# range（0， 5） 是[0, 1, 2, 3, 4]没有5
for ds_current_row in range(2, ds_sheet.max_row + 1):
    for item in ds_sheet_target_items:
        # 内部是个公式，需要给转化成字符串"X86企业级"、"0"、"X86终端",load_workbook时候，已经做了只要数据的处理
        # 现在给强制转换成字符串
        ds_gx_str = str(ds_sheet['G' + str(ds_current_row)].value)
        # 判断相等使用==，比较的是字符串内部，可以正常匹配；使用is是内存地址，如果2个字符串内存地址不一样，False
        if ds_gx_str == item:
            print('G' + str(ds_current_row) + '值：' + ds_gx_str + ' is ' + item + '=' + str(ds_gx_str == item))
            # 将数据源这一行添加到filter_result_sheet中
            ds_current_row_data = [item.value for item in ds_sheet[ds_current_row]]
            result_work_sheet.append(ds_current_row_data)

print("结果 sheet 列数: {0}".format(result_work_sheet.max_column))
print("结果 sheet 行数: {0}".format(result_work_sheet.max_row))

# 保存表格数据
time_string = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
filter_item_name = ''
for item in ds_sheet_target_items:
    filter_item_name = filter_item_name + '_' + item
print('filter_item_name = ' + filter_item_name)
result_work_book.save(time_string + '，过滤词' + filter_item_name + '，工作簿：' + workbook_name + '.xlsx')
