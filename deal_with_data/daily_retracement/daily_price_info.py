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