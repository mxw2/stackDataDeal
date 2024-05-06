from datetime import datetime

class PriceInfo:
    def __init__(self, date, high_value, low_value):
        assert high_value > 0, "high_value must be greater than 0"
        assert low_value > 0, "low_value must be greater than 0"
        if isinstance(date, str):
            self.date = date
        elif isinstance(date, datetime):
            self.date = date.strftime("%Y-%m-%d")
        else:
            raise TypeError("我也不知道什么类型", type(date))

        self.high_value = high_value
        self.low_value = low_value
        # 保留2位小数
        self.average_price = round((high_value + low_value) * 0.5, 2)
        # 应该是负数百分比，小数（未变成百分比）
        self.loss_percent = 0.0
        # 下降了多少钱，是负数
        self.loss_money = 0.0
        # 暴跌日价格信息
        self.loss_price_info: PriceInfo = None
        self.scale = 100

    # 获取损失百分比字符串
    def loss_percent_str(self):
        return f"{self.loss_percent_expand_100()}%"

    # 对百分比扩大100倍
    def loss_percent_expand_100(self):
        return round(self.loss_percent * self.scale, 2)

    def debug_model(self, index):
        index_str = f"【压测】翻转后的index: {index},"
        date_str = f"日期: {self.date},"
        high_value_str = f"最大值: {self.high_value},"
        low_value_str = f"最小值: {self.low_value},"
        avg_value_str = f"平均值: {self.average_price},"
        # loss_date_str = f"糟糕日: {self.loss_price_info.date},"
        loss_date_str = ""
        loss_percent_str = f"当轮【极端】回撤: {self.loss_percent_str()}"
        print(index_str + date_str + high_value_str + low_value_str + avg_value_str + loss_date_str + loss_percent_str)

    # def log_bad_day(self):
    #     # 打印自己, 这一天发生了大事，导致巨亏
    #     date_str = f"日期: {self.date},"
    #     low_value_str = f"最小值: {self.low_value}(不一定是这天暴跌，这里存在最低点)"
    #     # loss_percent_str = f"今天下降: {self.loss_percent_str()},"
    #     print(date_str + low_value_str)