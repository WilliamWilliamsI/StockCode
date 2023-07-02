import generate_index as st
import pandas as pd
import numpy as np

weakBuy, Buy, strongBuy, weakSell, Sell, strongSell, Wait = 1, 2, 3, 4, 5, 6, 7
PSY_VR_Signal_Up = False
PSY_VR_Signal_Down = False
data = {
    "Date": [],  # 时间
    "Open": [],  # 开盘价
    "Close": [],  # 收盘价
    "High": [],  # 最高价
    "Low": [],  # 最低价
    "Change": [],  # 涨跌额
    "ChangePercent": [],  # 涨跌幅
    "Turnover": [],  # 成交量
    "Volume": [],  # 成交额
    # "TurnoverRate": [],  # 换手率

    "MACD_EMA26": [],  # 26日慢速移动平均
    "MACD_EMA12": [],  # 12日快速移动平均
    "MACD_DIF": [],  # 求算MACD中间值
    "MACD_DEA": [],  # 求算MACD中间值
    "MACD_MACD": [],  # MACE柱状图, 即MACD

    "DMI_TR": [],
    "DMI_DI1": [],
    "DMI_DI2": [],
    "DMI_ADX": [],
    "DMI_ADXR": [],

    "EXPMA_MA1": [],
    "EXPMA_MA2": [],
    "EXPMA_MA3": [],
    "EXPMA_MA4": [],

    "PSY_PSY": [],

    "KDJ_K": [],  # KDJ的K值
    "KDJ_D": [],  # KDJ的D值
    "KDJ_J": [],  # KDJ的J值

    "BOLL_Mid": [],  # 布林线中轨线
    "BOLL_Top": [],  # 布林线上轨线
    "BOLL_Bottom": [],  # 布林线下轨线

    "WR1": [],  # 威廉超买超卖指标(N = 6日)
    "WR2": [],  # 威廉超买超卖指标(N = 10日)

    "RSI_RSI6": [],
    "RSI_RSI12": [],
    "RSI_RSI24": [],
    "RSI_tmp1_1": [],
    "RSI_tmp1_2": [],
    "RSI_tmp2_1": [],
    "RSI_tmp2_2": [],
    "RSI_tmp3_1": [],
    "RSI_tmp3_2": [],

    "BIAS_BIAS1": [],  # 6日BIAS曲线
    "BIAS_BIAS2": [],  # 12日BIAS曲线
    "BIAS_BIAS3": [],  # 24日BIAS曲线

    "VR_VR": [],

    "ARBR_AR": [],
    "ARBR_BR": [],

    "ASI_SI": [],
    "ASI_ASI": [],
    "ASI_ASIT": [],

    # "OBV_OBV": [],
    # "OBV_MAOBV": [],

    # "SAR_SAR_bearish": [],
    # "SAR_SAR_bullish": [],
    # "SAR_AR": [],

    "MIKE_WR": [],
    "MIKE_MR": [],
    "MIKE_SR": [],
    "MIKE_WS": [],
    "MIKE_MS": [],
    "MIKE_SS": [],

    "ROC_ROC": [],  # 变动率指标
    "ROC_MAROC": [],  # 变动率指标移动平均值

    "MTM_MTM": [],   # MTM
    "MTM_MTMMA": []  # MTMMA
}


def onMACD(i):
    """ onMACD
    :param:
        i: index
    :return
        Buy / Sell / Wait信号
    借助MACD辅助决策
        MACD指标说明
        MACD指数平滑异同移动平均线为两条长、短的平滑平均线。
        其买卖原则为：
        1.DIFF、DEA均为正，DIFF向上突破DEA，买入信号参考。
        2.DIFF、DEA均为负，DIFF向下跌破DEA，卖出信号参考。
        3.DEA线与K线发生背离，行情可能出现反转信号。
        4.分析MACD柱状线，由红变绿(正变负)，卖出信号参考；由绿变红，买入信号参考。
    """
    # todo DEA线与K线发生背离，行情可能出现反转信号。
    if i-1 >= 0:
        if data["MACD_DIF"][i] > data["MACD_DEA"][i] > 0 and data["MACD_DIF"][i-1] <= data["MACD_DEA"][i-1]:
            return Buy
        elif data["MACD_DIF"][i] < data["MACD_DEA"][i] < 0 and data["MACD_DIF"][i-1] >= data["MACD_DEA"][i-1]:
            return Sell
        if data["MACD_MACD"][i] < 0 and data["MACD_MACD"][i-1] > 0:
            return Sell
        elif data["MACD_MACD"][i] > 0 and data["MACD_MACD"][i-1] < 0:
            return Buy
        return Wait
    return Wait


def onDMI(i):
    """ onDMI
    :param:
        i: index
    :return
        Buy / Sell / Wait信号
    指示投资人避免在盘整的市场中交易，一旦市场变得有利润时，DMI立刻引导投资人进场，并且在适当时机退场。
    买卖原则：
        1、+DI上交叉-DI时，可参考做买。
        2、+DI下交叉-DI时，可参考做卖。
        3、ADX于50以上向下转折时，对表市场趋势终了。
        4、当ADX滑落至+DI之下时，不宜进场交易。
        5、当ADXR介于20-25时，宜采用TBP及CDP中之反应秘诀为交易参考。
    """
    if i-1 >= 0:
        if data["DMI_DI1"][i-1] < data["DMI_DI1"][i-1] and data["DMI_DI1"][i] > data["DMI_DI2"]:
            return Buy
        elif data["DMI_DI1"][i-1] > data["DMI_DI2"][i-1] and data["DMI_DI1"][i] < data["DMI_DI2"][i]:
            return Sell
        return Wait
    else:
        return Wait


def onEXPMA(i):
    """EXPMA
    :param:
        i: index
    :return
        Buy / Sell / Wait信号
    EXPMA
        为因移动平均线被视为落后指标的缺失而发展出来的，为解
    决一旦价格已脱离均线差值扩大，而平均线未能立即反应，EXPMA
    可以减少类似缺点。
    买卖原则：
    1、EXPMA译为指数平均线，修正移动平均线较股价落后的缺点，
       本指标随股价波动反应快速，用法与移动平均线相同。
    """
    pass


def onPSY_VR(i):
    """ 
        PSY
        1.PSY>75，形成Ｍ头时，股价容易遭遇压力；
        2.PSY<25，形成Ｗ底时，股价容易获得支撑；
        3.PSY 与VR 指标属一组指标群，须互相搭配使用。

        VR: VR实质为成交量的强弱指标，运用在过热之市场及低迷的盘局中，进一步辨认头部及底部的形成，有相当的作用，VR和PSY配合使用。
        买卖原则:
            1  VR下跌致40%以下时，市场极易形成底部。
            2  VR值一般分布在150%左右最多，一旦越过250%市场极易产生一段多头行情。
            3  VR超过350%以上，应有高档危机意识，随时注意反转之可能，可配合CR及PSY使用。
            4  VR运用在寻找底部时比较可靠，确认头部时，宜多配合其他指标使用。

        (1) 在PSY和VR组合中，投资者可观察PSY技术图形白线（PSY）处于25下方时，股价大概率会受到相对支撑区间，表现出市场超卖信号。
        加上VR技术图形中，当白线（VR）数值低于50以下，表示市场资金投入处于低位区间，市场成交低迷，后续图形出现白线（VR）出现转向上升，
        投资者可做为中期买入机会信号参考。
        (2) 当PSY技术图形白线（PSY）处于75上方时，股价大概率会受到相对压力区间，表现出市场超买信号。
        加上VR技术图形中，白线（VR）数值高于350以上，表示市场资金投入处于高位区间，市场成交过热，
        后续图形出现白线（VR）延续性减弱下降，投资者可做为短期卖出风险信号参考。
    """
    global PSY_VR_Signal_Up
    global PSY_VR_Signal_Down
    if data["PSY_PSY"][i] <= 25 and data["VR_VR"][i] <= 50:
        PSY_VR_Signal_Up = True
    if PSY_VR_Signal_Up and \
            data["VR_VR"][i] > data["VR_VR"][i-1] and data["VR_VR"][i] > data["VR_VR"][i-2] and \
            data["VR_VR"][i] > data["VR_VR"][i-3]:
        PSY_VR_Signal_Up = False
        return Buy
    if data["PSY_PSY"][i] >= 75 and data["VR_VR"][i] >= 350:
        PSY_VR_Signal_Down = True
    if PSY_VR_Signal_Down and \
            data["VR_VR"][i] < data["VR_VR"][i-1] and data["VR_VR"][i] < data["VR_VR"][i-2] and \
            data["VR_VR"][i] < data["VR_VR"][i-3]:
        PSY_VR_Signal_Down = False
        return Sell
    return Wait


def onKDJ(i):
    """
    KDJ指标
    reference: https://baike.baidu.com/item/KDJ%E6%8C%87%E6%A0%87/6328421?fromtitle=kdj&fromid=3423560&fr=aladdin#reference-[1]-348831-wrap
    指标说明
    KDJ，其综合动量观念、强弱指标及移动平均线的优点，早年应用在期货投资方面，功能颇为显著，目前为股市中最常被使用的指标之一。
    买卖原则
        1 K线由右边向下交叉D值做卖，K线由右边向上交叉D值做买。
        2 高档连续二次向下交叉确认跌势，低挡连续二次向上交叉确认涨势。
        3 D值<20超卖，D值>80超买，J>100超买，J<10超卖。
        4 KD值于50左右徘徊或交叉时，无意义。
        5 投机性太强的个股不适用。
        6 可观察KD值同股价的背离，以确认高低点。
    """
    # todo: 高档连续二次向下交叉确认跌势，低挡连续二次向上交叉确认涨势。
    if data["KDJ_K"][i-1] > data["KDJ_D"][i-1] and data["KDJ_K"][i] <= data["KDJ_D"][i]:
        return Sell
    elif data["KDJ_K"][i-1] < data["KDJ_D"][i-1] and data["KDJ_K"][i] >= data["KDJ_D"][i]:
        return Buy
    # # D值<20 超卖, 价格后期将会有上升, 可考虑买入
    # if data["KDJ_D"][i] < 20:
    #     return weakBuy
    # # D值>80 超买, 价格后期将会有下降
    # if data["KDJ_D"][i] > 80:
    #     return weakSell
    # # J>100 超买
    # if data["KDJ_J"][i] > 100:
    #     return weakSell
    # # J<10 超卖
    # if data["KDJ_J"][i] < 10:
    #     return weakBuy
    return Wait


def onWRShort(i):
    """ WR 威廉超买超卖指标(短期, 关注WR2)
    WR1, WR2 reference: https://zhidao.baidu.com/question/1183642184579614819.html
    用法：
        1.低于20，可能超买见顶，可考虑卖出
        2.高于80，可能超卖见底，可考虑买进
        3.与RSI、MTM指标配合使用，效果更好
    """
    if data["WR2"][i] < 20:
        return Sell
    if data["WR2"][i] > 80:
        return Buy
    return Wait


# WR: 威廉超买超卖指标(长期, 关注WR1)
def onWRLong(i):
    if data["WR1"][i] < 20:
        return Sell
    if data["WR1"][i] > 80:
        return Buy
    return Wait


def onRSI(i):
    """ RSI 指标: 
    RSI指标
    RSI的基本原理是在一个正常的股市中，多空买卖双方的力道必须得到均衡，股价才能稳定;而RSI是对于固定期间内，股价上涨总幅度平均值占总幅度平均值的比例。
        1 RSI值于0-100之间呈常态分配，当6日RSI值为80‰以上时，股市呈超买现象，若出现M头，市场风险较大；当6日RSI值在20‰以下时，股市呈超卖现象，若出现W头，市场机会增大。
        2 RSI一般选用6日、12日、24日作为参考基期，基期越长越有趋势性(慢速RSI)，基期越短越有敏感性，(快速RSI)。当快速RSI由下往上突破慢速RSI时，机会增大；当快速RSI由上而下跌破慢速RSI时，风险增大。
    """
    if data["RSI_RSI6"][i] > 80:
        return Sell
    if data["RSI_RSI6"][i] < 20:
        return Buy
    return Wait


def onBIAS(i):
    """ 乖离率 BIAS(6, 12, 24) 
    reference: https://zhuanlan.zhihu.com/p/38548544
        (1) 在弱势市场上,股价与６日移动平均线的乖离率达到＋６%以上时,为超买现象,是卖出时机；当其达到-６%以下时为超卖现象,是买入时机。而在强势市场上,
            股价与６日移动平均线乖离率达+８%以上时为超买现象,是卖出时机；当其达到-３%以下时为超卖现象,是买入时机。
        (2) 在弱势市场上,股价与１２日移动平均线的乖离率达到＋５%以上时为超买现象,是卖出时机,当其达到-５%以下时为超卖现象,是买入时机。而在强势市场上,
            股价与１２日移动平均线的乖离率达到+６%以上时为超买现象,是卖出时机,当其达到-４%以下时为超卖现象,是买入时机。
        (3) 对于个别股票,由于受多空双方激战的影响,股价和各种平均践的乖离率容易偏高,因此运用中要随之而变。
        (4) 当股价与平均线之间的乖离率达到最大百分比时,就会向零值逼近,有时也会低于零或高干零,这都属于正常现象。
        (5) 多头市场的暴涨和空头市场的暴跌,都会使乖离率达到意想不到的百分比值,但出现的次数极少,而且持续时间也很短,因此可以将其看作一种例外情形。
        (6) 在大势上涨的行情中,如果出现负乖离率,就可以在股价下跌时买进,此时损失的风险较小。BIAS负乖离率愈大，则空头回补的可能性愈高，如所示，
            短中长期各负乖离率均在-8%一下，这是买入的信号。
    实操: 未分析强势/弱势市场, 直接取最值, 减少交易次数
    """
    if data["BIAS_BIAS1"][i] > 6:
        return Sell
    if data["BIAS_BIAS1"][i] < -6:
        return Buy
    if data["BIAS_BIAS2"][i] > 6:
        return Sell
    if data["BIAS_BIAS2"][i] < -5:
        return Buy
    return Wait


# todo: ASI


# todo: MIKE

def onARBR(i):
    """ ARBR指标 
    reference: https://baike.baidu.com/item/arbr%E6%8C%87%E6%A0%87/2042243
    一、AR指标的单独研判
        1、AR值以100为强弱买卖气势的均衡状态，其值在上下20之间。亦即当AR值在80——120之间时，属于盘整行情，股价走势平稳，不会出现大幅上升或下降。
        2、AR值走高时表示行情活跃，人气旺盛，而过高则意味着股价已进入高价区，应随时卖出股票。在实际走势中，AR值的高度没有具体标准，一般情况下AR值大于180时（有的设定为150），预示着股价可能随时会大幅回落下跌，应及时卖出股票。
        3、AR值走低时表示行情萎靡不振，市场上人气衰退，而过低时则意味着股价可能已跌入低谷，随时可能反弹。一般情况下AR值小于40（有的设定为50）时，预示着股价已严重超卖，可考虑逢低介入。
        4、同大多数技术指标一样，AR指标也有领先股价到达峰顶和谷底的功能。当AR到达顶峰并回头时，如果股价还在上涨就应考虑卖出股票，获利了结；如果AR到达低谷后回头向上时，而股价还在继续下跌，就应考虑逢低买入股票。
    二、 BR指标的单独研判
        1、BR值为100时也表示买卖意愿的强弱呈平衡状态。
        2、BR值的波动比AR值敏感。当BR值介于70—150之间（有的设定为80—180）波动时，属于盘整行情，投资者应以观望为主。
        3、当BR值大于300（有的设定为400）时，表示股价进入高价区，可能随时回档下跌，应择机抛出。
        4、当BR值小于30（有的设定为40）时，表示股价已经严重超跌，可能随时会反弹向上，应逢低买入股票。
    三、 AR、BR指标的配合使用
        1、一般情况下，AR可单独使用，而BR则需与AR配合使用才能发挥BR的功能。
        2、AR和BR同时从低位向上攀升，表明市场上人气开始积聚，多头力量开始占优势，股价将继续上涨，投资者可及时买入或持筹待涨。
        3、当AR和BR从底部上扬一段时间后，到达一定高位并停滞不涨或开始掉头时，意味着股价已到达高位，持股者应注意及时获利了结。
        4、BR从高位回跌，跌幅达1/2时，若AR没有警戒信号出现，表明股价是上升途中的正常回调整理，投资者可逢低买入。
        5、当BR急速上升，而AR却盘整或小幅回档时，应逢高出货。
    """
    # todo: AR、BR指标的配合使用

    if data["ARBR_AR"][i] > 180:
        return Sell
    if data["ARBR_AR"][i] < 40:
        return Buy
    if data["ARBR_BR"][i] > 300:
        return Sell
    if data["ARBR_BR"][i] < 30:
        return Buy
    return Wait


def onBULL(i):
    """ BOLL: 布林线
    reference: https://www.zhihu.com/question/21094650?sort=created
    1.当股价穿越上限压力线（上轨）时，是卖出信号；
    2.当股价穿越下限支撑线（下轨）时，是买入信号；
    3.当股价处于上轨和中轨之间运行，表示股价目前处于多头市场；
    4.当股价处于下轨和中轨之间运行，表示股价目标处于空头市场；
    5.当股价由下向上穿越中轨时，为加码信号；
    6.当股价上向下穿越中轨时，为卖出信号。
    """
    if data["Close"][i-1] < data["BOLL_Top"][i-1] and data["Close"][i] > data["BOLL_Top"][i]:
        return Sell
    if data["Close"][i-1] > data["BOLL_Bottom"][i-1] and data["Close"][i] < data["BOLL_Bottom"][i]:
        return Buy
    if data["Close"][i-1] < data["BOLL_Mid"][i-1] and data["Close"][i] > data["BOLL_Mid"][i]:
        return Buy
    if data["Close"][i-1] > data["BOLL_Mid"][i-1] and data["Close"][i] < data["BOLL_Mid"][i]:
        return Sell
    return Wait


# 无投资经验投资者
# buy_rate: 股价上升时的
def withoutKnowledge(i, buy_rate=0.0001, sell_rate=0.0001):
    if data["Close"][i] > data["Close"][i-1] and (data["Close"][i] - data["Close"][i-1])/data["Close"][i-1] >= buy_rate:
        return Buy
    if data["Close"][i] < data["Close"][i-1] and (data["Close"][i-1] - data["Close"][i])/data["Close"][i] >= sell_rate:
        return Sell
    return Wait


def generateSignal():
    # 准备数据
    st.generateIndex()
    # 获取数据
    global data
    data = st.data
    signal = []
    for i in range(1785, len(data["Date"])):
        # 单次的买卖信号
        signal_per_day = [data["Date"][i]]
        res_MACD = onMACD(i)
        res_DMI = onDMI(i)
        res_PSY_VR = onPSY_VR(i)
        res_onKDJ = onKDJ(i)
        res_onWRShort = onWRShort(i)
        res_onWRLong = onWRLong(i)
        res_RSI = onRSI(i)
        res_BIAS = onBIAS(i)
        res_ARBR = onARBR(i)
        res_BULL = onBULL(i)
        res_knowledge = withoutKnowledge(i, 0.01, 0.01)
        res = [res_MACD, res_DMI, res_PSY_VR, res_onKDJ, res_onWRShort, res_onWRLong, res_RSI, res_BIAS, res_ARBR, res_BULL, res_knowledge]
        for ele in res:
            if ele == Buy:
                signal_per_day.append("Buy")
            elif ele == Sell:
                signal_per_day.append("Sell")
            # elif ele == weakSell:
            #     signal.append([data["Date"][i], "weakSell"])
            # elif ele == weakBuy:
            #     signal.append([data["Date"][i], "weakBuy"])
            elif ele == Wait:
                signal_per_day.append("Wait")
        signal.append(signal_per_day)
    # df = pd.DataFrame(signal, columns=["Date", "MACD", "DMI", "PSY_VR", "KDJ", "WR(s)", "WR(l)", "RSI", "BIAS", "ARBR", "BULL", "NoKnowledge"])
    # df.to_excel("decision.xlsx")


if __name__ == "__main__":
    generateSignal()
