class DailyPriceInfo:
    def __init__(self, max_value, min_value):
        if not all(isinstance(value, float) for value in [max_value, min_value]):
            raise ValueError("Both max_value and min_value must be of type float.")

        self.max_value = max_value
        self.min_value = min_value
        # 保留2位小数
        self.average_price = round((max_value + min_value) * 0.5, 2)