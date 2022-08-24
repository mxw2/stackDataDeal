from enum import Enum, unique


@unique
class CalculateType(Enum):
    Year = 1
    OriginalData = 2  # 原始数据即可
    DivisionIncome = 3  # 除以营业收入
    BusinessCreateProfit = 4  # 经营活动产生利润 (income - business_cost - business_tax - selling_expenses - manage_expenses - develop_expenses)
    GrossMargin = 5  # (1 - 成本占比) * 100 = 一个整数


useful_years_default = 10
useful_years_6 = 6
useful_years_4 = 4

# 数据源中收入的row值
ds_income_row_index = 12


class ColumnModel:
    # useful_years 显示几年，默认10年
    def __init__(self, name, ds_row_index, text_color, calculate_type, useful_years):
        self.name = name
        self.ds_row_index = ds_row_index
        self.text_color = text_color
        self.calculate_type = calculate_type
        self.useful_years = useful_years


def want_to_deal_with_data_source():
    # 具体年数
    years = ColumnModel('年份', 5, None, CalculateType.Year, useful_years_default)

    # 营业收入与经营活动现金流入图 (单位:亿元
    income = ColumnModel('营业收入（亿元）', ds_income_row_index, None, CalculateType.OriginalData, useful_years_default)
    cash_flow = ColumnModel('经营活动现金流入（亿元）', 244, None, CalculateType.OriginalData, useful_years_default)

    # 净利润现金净流对比分析(单位:亿元)
    net_profit = ColumnModel('净利润', 62, None, CalculateType.OriginalData, useful_years_default)
    business_net_cash = ColumnModel('经营活动产生的现金流量净额', 275, None, CalculateType.OriginalData, useful_years_default)

    # 历年现金流情况(单位:亿元)
    business_net_cash_short = ColumnModel('经营活动净额(亿元)', 275, "ff0000", CalculateType.OriginalData, useful_years_6)
    net_money_in = ColumnModel('筹资净额（亿元）', 318, "ff0000", CalculateType.OriginalData, useful_years_6)
    net_money_out = ColumnModel('投资净额（亿元）', 297, "ff0000", CalculateType.OriginalData, useful_years_6)

    # 历年收入成本构成(%)
    business_cost = ColumnModel('营业成本 %', 19, None, CalculateType.DivisionIncome, useful_years_4)
    business_tax = ColumnModel('营业税金及附加 %', 22, None, CalculateType.DivisionIncome, useful_years_4)
    selling_expenses = ColumnModel('销售费用 %', 23, None, CalculateType.DivisionIncome, useful_years_4)
    manage_expenses = ColumnModel('管理费用 %', 24, None, CalculateType.DivisionIncome, useful_years_4)
    develop_expenses = ColumnModel('研发费用 %', 25, None, CalculateType.DivisionIncome, useful_years_4)
    business_profit_expenses = ColumnModel('经营活动产生利润 %', 0, None, CalculateType.BusinessCreateProfit, useful_years_4)

    # 历年毛利率(%)
    gross_margin = ColumnModel('销售毛利率 %', 0, None, CalculateType.GrossMargin, useful_years_4)

    # 主要资产周转率（单位：次）

    # 历年资产堆积图
    cash = ColumnModel('货币资金（亿元）', 86, None, CalculateType.OriginalData, useful_years_default)
    goods = ColumnModel('存货（亿元）', 112, None, CalculateType.OriginalData, useful_years_default)
    fixed_asset = ColumnModel('固定资产（亿元）', 113, None, CalculateType.OriginalData, useful_years_default)
    building_project = ColumnModel('在建工程（亿元）', 134, None, CalculateType.OriginalData, useful_years_default)

    # 历年负债堆积图

    return [years,
            income,
            cash_flow,
            net_profit,
            business_net_cash,
            business_net_cash_short,
            net_money_in,
            net_money_out,
            business_cost,
            business_tax,
            selling_expenses,
            manage_expenses,
            develop_expenses,
            business_profit_expenses,
            gross_margin,
            cash,
            goods,
            fixed_asset,
            building_project]

