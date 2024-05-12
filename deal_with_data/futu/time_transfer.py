from datetime import datetime

def time_transfer_string():
    # 获取当前时间
    current_time = datetime.now()

    # 格式化时间戳为字符串
    formatted_time = current_time.strftime('%y%m%d_%H%M')

    print("格式化的当前时间:", formatted_time)
    return formatted_time