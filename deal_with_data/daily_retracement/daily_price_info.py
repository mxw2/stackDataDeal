class DailyPriceInfo:
    def __init__(self, date, max_value, min_value):
        assert max_value > 0, "max_value must be greater than 0"
        assert min_value > 0, "min_value must be greater than 0"
        if isinstance(date, str):
            self.date = date
        elif isinstance(date, date):
            self.date = date.strftime("%Y-%m-%d")
        else:
            raise TypeError("我也不知道什么类型 = ", type(date))

        self.max_value = max_value
        self.min_value = min_value
        # 保留2位小数
        self.average_price = round((max_value + min_value) * 0.5, 2)
        # 应该是负数百分比
        self.retracement_percent = 0.0

    def debug_log(self, index):
        index_str = f"index: {index},"
        date_str = f"日期: {self.date},"
        max_value_str = f"最大值: {self.max_value},"
        min_value_str = f"最小值: {self.min_value},"
        avg_value_str = f"平均值: {self.average_price},"
        loss_percent_str = f"回撤百分比: {round(self.retracement_percent * 100, 2)}%,"
        print(index_str + date_str + max_value_str + min_value_str + avg_value_str + loss_percent_str)