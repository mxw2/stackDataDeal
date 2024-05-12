import futu.common.constant

from futu_account_info import *


# 行情对象
quote_context = OpenQuoteContext(host=FUTUOPEND_ADDRESS,
                                 port=FUTUOPEND_PORT)
# 交易对象，根据交易品种修改交易对象类型
trade_context = OpenSecTradeContext(filter_trdmarket=TRADING_MARKET,
                                    host=FUTUOPEND_ADDRESS,
                                    port=FUTUOPEND_PORT,
                                    security_firm=SecurityFirm.FUTUSECURITIES)
# 我要拼出来一张dataFrame
option_chain_data_frame = None

# 解锁交易
def unlock_trade():
    if TRADING_ENVIRONMENT == TrdEnv.REAL:
        ret, data = trade_context.unlock_trade(TRADING_PWD)
        if ret != RET_OK:
            print('解锁交易失败：', data)
            return False
        print('解锁交易成功！')
    return True


# 获取市场状态
def is_normal_trading_time(code):
    ret, data = quote_context.get_market_state([code])
    if ret != RET_OK:
        print('获取市场状态失败：', data)
        return False
    market_state = data['market_state'][0]
    '''
    MarketState.MORNING            港、A 股早盘
    MarketState.AFTERNOON          港、A 股下午盘，美股全天
    MarketState.FUTURE_DAY_OPEN    港、新、日期货日市开盘
    MarketState.FUTURE_OPEN        美期货开盘
    MarketState.FUTURE_BREAK_OVER  美期货休息后开盘
    MarketState.NIGHT_OPEN         港、新、日期货夜市开盘
    '''
    if market_state == MarketState.MORNING or \
                    market_state == MarketState.AFTERNOON or \
                    market_state == MarketState.FUTURE_DAY_OPEN  or \
                    market_state == MarketState.FUTURE_OPEN  or \
                    market_state == MarketState.FUTURE_BREAK_OVER  or \
                    market_state == MarketState.NIGHT_OPEN:
        return True
    print('现在不是持续交易时段。')
    return False


# 获取持仓数量
def get_holding_position(code):
    holding_position = 0
    ret, data = trade_context.position_list_query(code=code, trd_env=TRADING_ENVIRONMENT)
    if ret != RET_OK:
        print('获取持仓数据失败：', data)
        return None
    else:
        for qty in data['qty'].values.tolist():
            holding_position += qty
        print('【持仓状态】 {} 的持仓数量为：{}'.format(TRADING_SECURITY, holding_position))
    return holding_position


# 拉取 K 线，计算均线，判断多空
def calculate_bull_bear(code, fast_param, slow_param):
    if fast_param <= 0 or slow_param <= 0:
        return 0
    if fast_param > slow_param:
        return calculate_bull_bear(code, slow_param, fast_param)
    ret, data = quote_context.get_cur_kline(code=code, num=slow_param + 1, ktype=TRADING_PERIOD)
    if ret != RET_OK:
        print('获取K线失败：', data)
        return 0
    candlestick_list = data['close'].values.tolist()[::-1]
    fast_value = None
    slow_value = None
    if len(candlestick_list) > fast_param:
        fast_value = sum(candlestick_list[1: fast_param + 1]) / fast_param
    if len(candlestick_list) > slow_param:
        slow_value = sum(candlestick_list[1: slow_param + 1]) / slow_param
    if fast_value is None or slow_value is None:
        return 0
    return 1 if fast_value >= slow_value else -1


# 获取一档摆盘的 ask1 和 bid1
def get_ask_and_bid(code):
    ret, data = quote_context.get_order_book(code, num=1)
    if ret != RET_OK:
        print('获取摆盘数据失败：', data)
        return None, None
    return data['Ask'][0][0], data['Bid'][0][0]


# 开仓函数
def open_position(code):
    # 获取摆盘数据
    ask, bid = get_ask_and_bid(code)

    # 计算下单量
    open_quantity = calculate_quantity()

    # 判断购买力是否足够
    if is_valid_quantity(TRADING_SECURITY, open_quantity, ask):
        # 下单
        ret, data = trade_context.place_order(price=ask, qty=open_quantity, code=code, trd_side=TrdSide.BUY,
                                              order_type=OrderType.NORMAL, trd_env=TRADING_ENVIRONMENT,
                                              remark='moving_average_strategy')
        if ret != RET_OK:
            print('开仓失败：', data)
    else:
        print('下单数量超出最大可买数量。')


# 平仓函数
def close_position(code, quantity):
    # 获取摆盘数据
    ask, bid = get_ask_and_bid(code)

    # 检查平仓数量
    if quantity == 0:
        print('无效的下单数量。')
        return False

    # 平仓
    ret, data = trade_context.place_order(price=bid, qty=quantity, code=code, trd_side=TrdSide.SELL,
                   order_type=OrderType.NORMAL, trd_env=TRADING_ENVIRONMENT, remark='moving_average_strategy')
    if ret != RET_OK:
        print('平仓失败：', data)
        return False
    return True


# 计算下单数量
def calculate_quantity():
    price_quantity = 0
    # 使用最小交易量
    ret, data = quote_context.get_market_snapshot([TRADING_SECURITY])
    if ret != RET_OK:
        print('获取快照失败：', data)
        return price_quantity
    price_quantity = data['lot_size'][0]
    return price_quantity


# 判断购买力是否足够
def is_valid_quantity(code, quantity, price):
    ret, data = trade_context.acctradinginfo_query(order_type=OrderType.NORMAL,
                                                   code=code,
                                                   price=price,
                                                   trd_env=TRADING_ENVIRONMENT)
    if ret != RET_OK:
        print('获取最大可买可卖失败：', data)
        return False
    max_can_buy = data['max_cash_buy'][0]
    max_can_sell = data['max_sell_short'][0]
    if quantity > 0:
        return quantity < max_can_buy
    elif quantity < 0:
        return abs(quantity) < max_can_sell
    else:
        return False


# 展示订单回调
def show_order_status(data):
    order_status = data['order_status'][0]
    order_info = dict()
    order_info['代码'] = data['code'][0]
    order_info['价格'] = data['price'][0]
    order_info['方向'] = data['trd_side'][0]
    order_info['数量'] = data['qty'][0]
    print('【订单状态】', order_status, order_info)


############################ 迎新策略逻辑处理 ############################
# 一.订阅股票的行情
def subscribe_security():
    # 设置回调
    quote_context.set_handler(OnTickClass())
    trade_context.set_handler(OnOrderClass())
    trade_context.set_handler(OnFillClass())

    print('************  一.订阅股票的行情 ***********')
    # TICKER: 订阅标的合约的逐笔(当前股价)
    # ORDER_BOOK: 摆盘(这个求平均值，我认为更可靠)
    quote_context.subscribe(code_list=[TRADING_SECURITY],
                            subtype_list=[SubType.TICKER, SubType.ORDER_BOOK])


# 二.查看当前股票的期权到期日列表
def get_security_option_expiration_date():
    print('************  二.查看当前股票的期权到期日列表 ***********')
    # 我只请求一次即可，不怕时间限制 https://openapi.futunn.com/futu-api-doc/quote/get-option-expiration-date.html
    ret, data = quote_context.get_option_expiration_date(code=TRADING_SECURITY)
    if ret == RET_OK:
        print(data)
        print(data['strike_time'].values.tolist())  # 转为 list
        get_security_option_list(data)
    else:
        print('error:', data)

# 三.股票的期权每一个到期日的期权链
def get_security_option_list(data):
    global option_chain_data_frame
    print('************  三.股票的期权每一个到期日的期权链 ***********')
    # 三.查看当前期权到期日列表对应的每一个期权连
    # 通过标的股票查询期权链。此接口仅返回期权链的静态信息，如需获取报价或摆盘等动态信息，
    # 请用此接口返回的股票代码，自行 订阅 所需要的类型。
    filter1 = OptionDataFilter()
    # 不要轻易设置delta_min = 0.01, 会导致-0.01的put无法被发现
    filter1.delta_min = -0.6
    filter1.delta_max = -0.002

    expiration_date_list = data['strike_time'].values.tolist()
    for date in expiration_date_list:
        ret2, data2 = quote_context.get_option_chain(code=TRADING_SECURITY,
                                                     start=date,
                                                     end=date,
                                                     option_type=OptionType.PUT,
                                                     option_cond_type=OptionCondType.OUTSIDE, # 价外期权
                                                     data_filter=filter1)
        if ret2 == RET_OK:
            print('-----------↓↓↓↓↓↓↓↓↓↓↓↓期权：' + data2['code'][0] + '↓↓↓↓↓↓↓↓↓↓↓↓-----------')  # 取第一条的股票代码
            # print(data2)
            if option_chain_data_frame is None:
                option_chain_data_frame = data2
            else:
                option_chain_data_frame = pd.concat([option_chain_data_frame, data2], ignore_index=True, axis=0)
            print(data2['code'].values.tolist())  # 转为 list
        else:
            print('error:', data2)
        # 这3s会导致股价和期权价格不对等！！！，需要想办法解决它
        time.sleep(3)

    print('所有日期的期权链打印完毕哈')
    print(option_chain_data_frame)


# 四.对期权链进行订阅，获取他们的股价 & 期权价格
def subscribe_option_chain():
    print('************  四.对期权链进行订阅，获取他们的股价 & 期权价格 ***********')
    print(option_chain_data_frame['code'].values.tolist())


############################ 填充以下函数来完成您的策略 ############################
# 策略启动时运行一次，用于初始化策略
def on_init():
    # 解锁交易（如果是模拟交易则不需要解锁）
    if not unlock_trade():
        return False
    print('************  策略开始运行 ***********')
    return True


# 每个 tick 运行一次，可将策略的主要逻辑写在此处
def on_tick(data):
    print('TRADING_SECURITY ' + TRADING_SECURITY)


# 每次产生一根新的 K 线运行一次，可将策略的主要逻辑写在此处
# def on_bar_open():
#     # 打印分隔线
#     print('*************************************')
#
#     # 只在常规交易时段交易
#     if not is_normal_trading_time(TRADING_SECURITY):
#         return
#
#     # 获取 K 线，计算均线，判断多空
#     bull_or_bear = calculate_bull_bear(TRADING_SECURITY, FAST_MOVING_AVERAGE, SLOW_MOVING_AVERAGE)
#
#     # 获取持仓数量
#     holding_position = get_holding_position(TRADING_SECURITY)
#
#     # 下单判断
#     if holding_position == 0:
#         if bull_or_bear == 1:
#             print('【操作信号】 做多信号，建立多单。')
#             open_position(TRADING_SECURITY)
#         else:
#             print('【操作信号】 做空信号，不开空单。')
#     elif holding_position > 0:
#         if bull_or_bear == -1:
#             print('【操作信号】 做空信号，平掉持仓。')
#             close_position(TRADING_SECURITY, holding_position)
#         else:
#             print('【操作信号】 做多信号，无需加仓。')


# 委托成交有变化时运行一次
def on_fill(data):
    if data['code'][0] == TRADING_SECURITY:
        # order_status = data['order_status'][0]
        # order_info = dict()
        # order_info['代码'] = data['code'][0]
        # order_info['价格'] = data['price'][0]
        # order_info['方向'] = data['trd_side'][0]
        # order_info['数量'] = data['qty'][0]
        # print('【订单状态】', order_status, order_info)
        print(TRADING_SECURITY)


# 订单状态有变化时运行一次
def on_order_status(data):
    if data['code'][0] == TRADING_SECURITY:
        show_order_status(data)


################################ 框架实现部分，可忽略不看 ###############################
class OnTickClass(TickerHandlerBase):
    # marsxw : 异步完成每一笔交易的回调，遇到流通性不好的会有较大差距
    # 目前可以在这里填充下股票价格 & 期权价格，但是期权价格建议采用平均值更好点
    def on_recv_rsp(self, rsp_pb):
        print('************  股票的每一次成交回调 ***********')
        ret, data = super(OnTickClass, self).on_recv_rsp(rsp_pb)
        if ret == RET_OK:
            on_tick(data)


# class OnBarClass(CurKlineHandlerBase):
#     last_time = None
#     def on_recv_rsp(self, rsp_pb):
        # 我不关注K线
        # ret_code, data = super(OnBarClass, self).on_recv_rsp(rsp_pb)
        # if ret_code == RET_OK:
        #     cur_time = data['time_key'][0]
        #     if cur_time != self.last_time and data['k_type'][0] == TRADING_PERIOD:
        #         if self.last_time is not None:
        #             on_bar_open()
        #         self.last_time = cur_time


class OnOrderClass(TradeOrderHandlerBase):
    def on_recv_rsp(self, rsp_pb):
        print('************  OnOrderClass ***********')
        ret, data = super(OnOrderClass, self).on_recv_rsp(rsp_pb)
        if ret == RET_OK:
            on_order_status(data)


class OnFillClass(TradeDealHandlerBase):
    def on_recv_rsp(self, rsp_pb):
        print('************  OnFillClass ***********')
        ret, data = super(OnFillClass, self).on_recv_rsp(rsp_pb)
        if ret == RET_OK:
            on_fill(data)


# 主函数
if __name__ == '__main__':
    # 初始化策略
    if not on_init():
        print('策略初始化失败，脚本退出！')
        quote_context.close()
        trade_context.close()
    else:
        subscribe_security()
        get_security_option_expiration_date()



