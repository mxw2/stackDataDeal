from enum import Enum, unique


@unique
class CalculateType(Enum):
    Year = 1
    OriginalData = 2  # 原始数据即可
    DivisionIncome = 3  # 除以营业收入


class ColumnModel:
    def __init__(self, name, ds_row_index, text_color, calculate_type):
        self.name = name
        self.ds_row_index = ds_row_index
        self.text_color = text_color
        self.calculate_type = calculate_type


def want_to_deal_with_data_source():
    years = ColumnModel('年份', 5, None, CalculateType.Year)
    income = ColumnModel('营业收入（亿元）', 12, None, CalculateType.OriginalData)
    cash_flow = ColumnModel('经营活动现金流入（亿元）', 244, None, CalculateType.OriginalData)
    net_profit = ColumnModel('净利润', 62, None, CalculateType.OriginalData)
    business_net_cash = ColumnModel('经营活动产生的现金流量净额', 275, None, CalculateType.OriginalData)
    net_money_in = ColumnModel('筹资净额（亿元）', 318, None, CalculateType.OriginalData)
    net_money_out = ColumnModel('投资净额（亿元）', 297, None, CalculateType.OriginalData)
    business_money_in = ColumnModel('营业成本%', 19, None, CalculateType.DivisionIncome)

    return [years,
            income,
            cash_flow,
            net_profit,
            business_net_cash,
            net_money_in,
            net_money_out,
            business_money_in]

