from datetime import datetime

class PriceInfo:
    def __init__(self, date, max_value, min_value):
        assert max_value > 0, "max_value must be greater than 0"
        assert min_value > 0, "min_value must be greater than 0"
        if isinstance(date, str):
            self.date = date
        elif isinstance(date, datetime):
            self.date = date.strftime("%Y-%m-%d")
        else:
            raise TypeError("我也不知道什么类型", type(date))

        self.max_value = max_value
        self.min_value = min_value
        # 保留2位小数
        self.average_price = round((max_value + min_value) * 0.5, 2)
        # 应该是负数百分比，小数（未变成百分比）
        self.loss_percent = 0.0
        # 下降了多少钱，是负数
        self.loss_money = 0.0
        # 暴跌日价格信息
        self.loss_price_info: PriceInfo = None

    # 获取损失百分比字符串，会对
    def loss_percent_str(self):
        return f"{round(self.loss_percent * 100, 2)}%"

    def debug_model(self, index):
        index_str = f"翻转后的index: {index},"
        date_str = f"日期: {self.date},"
        max_value_str = f"最大值: {self.max_value},"
        min_value_str = f"最小值: {self.min_value},"
        avg_value_str = f"平均值: {self.average_price},"
        # today_loss_percent_str = f"当天涨跌: {self.}"
        loss_percent_str = f"当轮回撤百分比: {self.loss_percent_str()},"
        print(index_str + date_str + max_value_str + min_value_str + avg_value_str + loss_percent_str)