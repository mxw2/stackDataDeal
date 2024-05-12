from futu import *

############################ 全局变量设置 ############################
FUTUOPEND_ADDRESS = '127.0.0.1'  # OpenD 监听地址
FUTUOPEND_PORT = 11111  # OpenD 监听端口

TRADING_ENVIRONMENT = TrdEnv.REAL  # 交易环境：真实 / 模拟
TRADING_MARKET = TrdMarket.HK  # 交易市场权限，用于筛选对应交易市场权限的账户
TRADING_PWD = '521647'  # 交易密码，用于解锁交易
# TRADING_PERIOD = KLType.K_1M  # 信号 K 线周期
TRADING_SECURITY = 'HK.03690'  # 交易标的'HK.TCH240530P250000'#
# FAST_MOVING_AVERAGE = 1  # 均线快线的周期
# SLOW_MOVING_AVERAGE = 3  # 均线慢线的周期


# 行情对象
quote_context = OpenQuoteContext(host=FUTUOPEND_ADDRESS,
                                 port=FUTUOPEND_PORT)
# 交易对象，根据交易品种修改交易对象类型
trade_context = OpenSecTradeContext(filter_trdmarket=TRADING_MARKET,
                                    host=FUTUOPEND_ADDRESS,
                                    port=FUTUOPEND_PORT,
                                    security_firm=SecurityFirm.FUTUSECURITIES)