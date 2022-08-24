from enum import Enum, unique


@unique
class CalculateType(Enum):
    Year = 1
    OriginalData = 2  # 原始数据即可
    DivisionIncome = 3  # 除以营业收入
    BusinessCreateProfit = 4  # 经营活动产生利润 (income - business_cost - business_tax - selling_expenses - manage_expenses - develop_expenses)
    GrossMargin = 5  # (1 - 成本占比) * 100 = 一个整数
    CommonTurnoverRate = 6  # 常用的周转率 B12/((B153+A153)/2) : B12 销售收入总额
    GoodsTurnoverRate = 7  # B19/((B112+A112)/2) : B19 成本总额


useful_years_default = 10
useful_years_6 = 6
useful_years_4 = 4

# 数据源中收入的row值
ds_income_row = 12
ds_business_cost_row = 19
ds_business_tax_row = 22
ds_selling_expenses_row = 23
ds_manage_expenses_row = 24
ds_develop_expenses_row = 25

# 提供给外界用的可变数组
models = []


class ColumnModel:
    # useful_years 显示几年，默认10年
    def __init__(self, name, ds_row_index, text_color, calculate_type, useful_years):
        self.name = name
        self.ds_row_index = ds_row_index
        self.text_color = text_color
        self.calculate_type = calculate_type
        self.useful_years = useful_years


def create_column_model(name, ds_row_index, text_color, calculate_type, useful_years):
    model = ColumnModel(name, ds_row_index, text_color, calculate_type, useful_years)
    models.append(model)


def want_to_deal_with_data_source():
    # 具体年数
    create_column_model('年份', 5, None, CalculateType.Year, useful_years_default)

    # 营业收入与经营活动现金流入图 (单位:亿元
    create_column_model('营业收入（亿元）', ds_income_row, "00ff00", CalculateType.OriginalData, useful_years_default)
    create_column_model('经营活动现金流入（亿元）', 244, "00ff00", CalculateType.OriginalData, useful_years_default)

    # 净利润现金净流对比分析(单位:亿元)
    create_column_model('净利润', 62, "0000ff", CalculateType.OriginalData, useful_years_default)
    create_column_model('经营活动产生的现金流量净额', 275, "0000ff", CalculateType.OriginalData, useful_years_default)

    # 历年现金流情况(单位:亿元)
    create_column_model('经营活动净额(亿元)', 275, "ff0000", CalculateType.OriginalData, useful_years_6)
    create_column_model('筹资净额（亿元）', 318, "ff0000", CalculateType.OriginalData, useful_years_6)
    create_column_model('投资净额（亿元）', 297, "ff0000", CalculateType.OriginalData, useful_years_6)

    # 历年收入成本构成(%)
    create_column_model('营业成本 %', ds_business_cost_row, "aaff00", CalculateType.DivisionIncome, useful_years_4)
    create_column_model('营业税金及附加 %', ds_business_tax_row, "aaff00", CalculateType.DivisionIncome, useful_years_4)
    create_column_model('销售费用 %', ds_selling_expenses_row, "aaff00", CalculateType.DivisionIncome, useful_years_4)
    create_column_model('管理费用 %', ds_manage_expenses_row, "aaff00", CalculateType.DivisionIncome, useful_years_4)
    create_column_model('研发费用 %', ds_develop_expenses_row, "aaff00", CalculateType.DivisionIncome, useful_years_4)
    create_column_model('经营活动产生利润 %', 0, "aaff00", CalculateType.BusinessCreateProfit, useful_years_4)

    # 历年毛利率(%)
    create_column_model('销售毛利率 %', 0, "00ffaa", CalculateType.GrossMargin, useful_years_4)

    # 主要资产周转率（单位：次）
    create_column_model('总资产周转率（亿元）', 153, "00ccaa", CalculateType.CommonTurnoverRate, useful_years_4)
    create_column_model('应收账款周转率（亿元）', 95, "00ccaa", CalculateType.CommonTurnoverRate, useful_years_4)
    create_column_model('存货周转率（亿元）', 112, "00ccaa", CalculateType.GoodsTurnoverRate, useful_years_4)
    create_column_model('固定资产周转率（亿元）', 133, "00ccaa", CalculateType.CommonTurnoverRate, useful_years_4)

    # 历年资产堆积图
    create_column_model('货币资金（亿元）', 86, "bbccaa", CalculateType.OriginalData, useful_years_default)
    create_column_model('存货（亿元）', 112, "bbccaa", CalculateType.OriginalData, useful_years_default)
    create_column_model('应收票据及应收账款（亿元）', 95, "bbccaa", CalculateType.OriginalData, useful_years_default)
    create_column_model('固定资产（亿元）', 113, "bbccaa", CalculateType.OriginalData, useful_years_default)
    create_column_model('在建工程（亿元）', 134, "bbccaa", CalculateType.OriginalData, useful_years_default)

    # 历年负债堆积图
    create_column_model('应付票据及应付账款（亿元）', 164, "bb2233", CalculateType.OriginalData, useful_years_default)
    create_column_model('其他应付款（亿元）', 173, "bb2233", CalculateType.OriginalData, useful_years_default)

    return models

