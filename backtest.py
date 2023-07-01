import copy
import random
import math
import pandas
import xlrd
import datetime
import matplotlib.dates as mdates
from mpl_finance import candlestick_ohlc
import matplotlib.ticker as ticker
import matplotlib.font_manager as fm
import generate_index as gi
import seaborn as sns
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import gridspec
import pymysql

Buy, Sell, Wait = "Buy", "Sell", "Wait"


class Stock:
    def __init__(self):
        self.data = {
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
            "PSY_PSYMA": [],

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

            "MIKE_WR": [],
            "MIKE_MR": [],
            "MIKE_SR": [],
            "MIKE_WS": [],
            "MIKE_MS": [],
            "MIKE_SS": [],

            "ROC_ROC": [],  # 变动率指标
            "ROC_MAROC": [],  # 变动率指标移动平均值

            "MTM_MTM": [],  # MTM
            "MTM_MTMMA": []  # MTMMA
        }


stockInfo = []  # 所有股票的数据信息


def first_Row(workbook, stockIndex):
    """
    返回要处理的首行
    如果无信息, 则返回1
    """
    if 2020 in workbook.sheet_by_name(stockIndex).col_values(6):
        row = workbook.sheet_by_name(stockIndex).col_values(6).index(2020)
        if row > 50:
            return row
        else:
            return 50
    elif 2021 in workbook.sheet_by_name(stockIndex).col_values(6):
        return 50
    elif 2022 in workbook.sheet_by_name(stockIndex).col_values(6):
        return 50
    else:
        return -1


def onMACD(i, stockIndex):
    data = stockInfo[stockIndex].data
    # print(data["MACD_DIF"][i], data["MACD_DEA"][i], data["MACD_DIF"][i-1], data["MACD_DEA"][i-1])
    # print(data["MACD_DIF"][i] < data["MACD_DEA"][i], data["MACD_DIF"][i-1] > data["MACD_DEA"][i-1])
    # print(i, data["MACD_MACD"][i], data["MACD_MACD"][i-1], data["MACD_MACD"][i] < 0 and data["MACD_MACD"][i - 1] > 0)
    if i - 1 >= 0:
        if data["MACD_DIF"][i] > data["MACD_DEA"][i] and data["MACD_DIF"][i - 1] < data["MACD_DEA"][i - 1]:
            return Buy
        if data["MACD_DIF"][i] < data["MACD_DEA"][i] and data["MACD_DIF"][i - 1] > data["MACD_DEA"][i - 1]:
            return Sell
        if data["MACD_MACD"][i] < 0 and data["MACD_MACD"][i - 1] > 0:
            return Sell
        if data["MACD_MACD"][i] > 0 and data["MACD_MACD"][i - 1] < 0:
            return Buy
        return Wait
    return Wait


def onDMI(i, stockIndex):
    data = stockInfo[stockIndex].data
    if i - 1 >= 0:
        if data["DMI_DI1"][i - 1] < data["DMI_DI2"][i - 1] and data["DMI_DI1"][i] > data["DMI_DI2"][i]:
            return Buy
        elif data["DMI_DI1"][i - 1] > data["DMI_DI2"][i - 1] and data["DMI_DI1"][i] < data["DMI_DI2"][i]:
            return Sell
        return Wait
    else:
        return Wait


def onPSY(i, stockIndex):
    data = stockInfo[stockIndex].data
    """https://baike.baidu.com/item/PSY%E6%8C%87%E6%A0%87/3083493?fr=aladdin"""
    if data["PSY_PSY"][i] > 90:
        return Sell
    if data["PSY_PSY"][i] < 10:
        return Buy
    if i >= 1:
        # 当PSY曲线和PSYMA曲线同时向上运行时，为买入时机
        if data["PSY_PSYMA"][i] > data["PSY_PSYMA"][i - 1] and data["PSY_PSY"][i] > data["PSY_PSY"][i - 1]:
            return Buy
        # 当PSY曲线与PSYMA曲线同时向下运行时，为卖出时机
        if data["PSY_PSYMA"][i] < data["PSY_PSYMA"][i - 1] and data["PSY_PSY"][i] < data["PSY_PSY"][i - 1]:
            return Sell
        # 当PSY曲线向上突破PSYMA曲线时，为买入时机
        if data["PSY_PSY"][i] > data["PSY_PSYMA"][i] and data["PSY_PSY"][i - 1] < data["PSY_PSYMA"][i - 1]:
            return Buy
        # 当PSY曲线向下跌破PSYMA曲线后，为卖出时机
        if data["PSY_PSY"][i] < data["PSY_PSYMA"][i] and data["PSY_PSY"][i - 1] > data["PSY_PSYMA"][i - 1]:
            return Sell
    return Wait


def onVR(i, stockIndex):
    data = stockInfo[stockIndex].data
    """https://baike.baidu.com/item/%E6%88%90%E4%BA%A4%E9%87%8F%E5%8F%98%E5%BC%82%E7%8E%87/1976493?fromtitle=VR%E6%8C
    %87%E6%A0%87&fromid=4672285&fr=aladdin """
    # VR<40，市场易形成底部，应积极买入
    if data["VR_VR"][i] < 40:
        return Buy
    # 当VR>450，市场交易过热，应卖出
    if data["VR_VR"][i] > 450:
        return Sell
    return Wait


def onKDJ(i, stockIndex):
    data = stockInfo[stockIndex].data
    """https://baike.baidu.com/item/KDJ%E6%8C%87%E6%A0%87/6328421?fr=aladdin"""
    if i >= 1:
        # D大于80时，行情呈现超买现象
        if data["KDJ_D"][i] > 80:
            return Sell
        # D小于20时，行情呈现超卖现象
        if data["KDJ_D"][i] < 20:
            return Buy
        # 上涨趋势中，K值大于D值，K线向上突破D线时，为买进信号
        if data["KDJ_K"][i - 1] > data["KDJ_D"][i - 1] and data["KDJ_K"][i] < data["KDJ_D"][i]:
            return Sell
        # 下跌趋势中，K值小于D值，K线向下跌破D线时，为卖出信号
        if data["KDJ_K"][i - 1] < data["KDJ_D"][i - 1] and data["KDJ_K"][i] > data["KDJ_D"][i]:
            return Buy
    return Wait


def onWR(i, stockIndex):
    data = stockInfo[stockIndex].data
    """https://baike.baidu.com/item/%E5%A8%81%E5%BB%89%E6%8C%87%E6%A0%87/2064411?fromtitle=wR%E6%8C%87%E6%A0%87
    &fromid=9464148&fr=aladdin """
    # 当W&R高于80，即处于超卖状态，行情即将见底，应当考虑买进
    if data["WR1"][i] < 20:
        return Sell
    # 当W&R低于20，即处于超买状态，行情即将见顶，应当考虑卖出
    if data["WR1"][i] > 80:
        return Buy
    if i - 1 >= 0:
        # 在W&R进入高位后，一般要回头，如果股价继续上升就产生了背离，是卖出信号。
        if data["WR1"][i - 1] > 70 and data["WR1"][i] < data["WR1"][i - 1] and data["Close"][i] > data["Close"][i - 1]:
            return Sell
        # 在W&R进入低位后，一般要反弹，如果股价继续下降就产生了背离。
        if data["WR1"][i - 1] < 30 and data["WR1"][i] > data["WR1"][i - 1] and data["Close"][i] < data["Close"][i - 1]:
            return Buy
    return Wait


def onRSI(i, stockIndex):
    data = stockInfo[stockIndex].data
    """https://baike.baidu.com/item/%E7%9B%B8%E5%AF%B9%E5%BC%BA%E5%BC%B1%E6%8C%87%E6%A0%87/6838822?fromtitle=RSI
    &fromid=6130115&fr=aladdin """
    """https://www.moomoo.com/sg/support/topic3_239?from_platform=4&platform_langArea=sg"""
    # 当RSI值超过80时，则表示整个市场力度过强，多方力量远大于空方力量，双方力量对比悬殊，多方大胜，市场处于超买状态，后续行情有可能出现回调或转势，此时，投资者可卖出股票。
    if data["RSI_RSI6"][i] > 80:
        return Sell
    # 当RSI值低于20时，则表示市场上卖盘多于买盘，空方力量强于多方力量，空方大举进攻后，市场下跌的幅度过大，已处于超卖状态，股价可能出现反弹或转势，投资者可适量建仓、买入股票。
    if data["RSI_RSI6"][i] < 20:
        return Buy
    return Wait


def onBIAS(i, stockIndex):
    data = stockInfo[stockIndex].data
    """https://www.moomoo.com/sg/support/topic3_245?lang=en-us&from_platform=4&platform_langArea=sg"""
    """https://baike.baidu.com/item/BIAS/9060339?fromModule=lemma-qiyi_sense-lemma"""
    # 6日平均值乖离：－3%是买进时机，+3．5是卖出时机；
    if data["BIAS_BIAS1"][i] > 3.5:
        return Sell
    if data["BIAS_BIAS1"][i] < -3:
        return Buy
    # 12日平均值乖离：－4．5%是买进时机，+5%是卖出时机；
    if data["BIAS_BIAS2"][i] > 5:
        return Sell
    if data["BIAS_BIAS2"][i] < -4.5:
        return Buy
    # 24日平均值乖离：－7%是买进时机，+8%是卖出时机；
    if data["BIAS_BIAS3"][i] > 8:
        return Sell
    if data["BIAS_BIAS3"][i] < -7:
        return Buy
    return Wait


def onARBR(i, stockIndex):
    data = stockInfo[stockIndex].data
    """https://www.moomoo.com/sg/support/topic3_157"""
    if data["ARBR_AR"][i] > 150:
        return Sell
    if data["ARBR_AR"][i] < 50:
        return Buy
    if data["ARBR_BR"][i] > 300:
        return Sell
    if data["ARBR_BR"][i] < 50:
        return Buy
    return Wait


def onBOLL(i, stockIndex):
    data = stockInfo[stockIndex].data
    if i - 1 >= 0:
        """https://www.moomoo.com/sg/support/topic3_241?from_platform=4&platform_langArea=sg"""
        if data["Close"][i - 1] < data["BOLL_Top"][i - 1] and data["Close"][i] > data["BOLL_Top"][i]:
            return Sell
        if data["Close"][i - 1] > data["BOLL_Bottom"][i - 1] and data["Close"][i] < data["BOLL_Bottom"][i]:
            return Buy
        if data["Close"][i - 1] < data["BOLL_Mid"][i - 1] and data["Close"][i] > data["BOLL_Mid"][i]:
            return Buy
        if data["Close"][i - 1] > data["BOLL_Mid"][i - 1] and data["Close"][i] < data["BOLL_Mid"][i]:
            return Sell
    return Wait


def onROC(i, stockIndex):
    data = stockInfo[stockIndex].data
    if i - 1 >= 0:
        if data["ROC_ROC"][i] < 0 and data["ROC_ROC"][i - 1] > 0:
            return Sell
        if data["ROC_ROC"][i] > 0 and data["ROC_ROC"][i - 1] < 0:
            return Buy
        if data["ROC_ROC"][i - 1] > data["ROC_MAROC"][i - 1] and data["ROC_ROC"][i] < data["ROC_MAROC"][i]:
            return Sell
        if data["ROC_ROC"][i - 1] < data["ROC_MAROC"][i - 1] and data["ROC_ROC"][i] > data["ROC_MAROC"][i]:
            return Buy
    return Wait


def onMTM(i, stockIndex):
    data = stockInfo[stockIndex].data
    """http://www.990755.com/stock/1725.html"""
    if i - 1 >= 0:
        if data["MTM_MTM"][i - 1] < data["MTM_MTMMA"][i - 1] and data["MTM_MTM"][i] > data["MTM_MTMMA"][i]:
            return Buy
        if data["MTM_MTM"][i - 1] > data["MTM_MTMMA"][i - 1] and data["MTM_MTM"][i] < data["MTM_MTMMA"][i]:
            return Sell
    return Wait


def onNULL(i, stockIndex):
    return Buy


# 追涨杀跌
def CUKD(i, stockIndex, buy_rate=0.03, sell_rate=0.03):
    data = stockInfo[stockIndex].data
    if i - 1 >= 0:
        # 今日较昨日增长超过3%, 跟随买入
        if data["Close"][i] > data["Close"][i - 1] and \
                (data["Close"][i] - data["Close"][i - 1]) / data["Close"][i - 1] >= buy_rate:
            return Buy
        if data["Close"][i] < data["Close"][i - 1] and \
                (data["Close"][i - 1] - data["Close"][i]) / data["Close"][i - 1] >= sell_rate:
            return Sell
    return Wait


# 高抛低吸
def HTLS(i, stockIndex, buy_rate=0.05, sell_rate=0.05):
    data = stockInfo[stockIndex].data
    if i - 5 >= 0:
        # 相对昨日的增长率超过5%, 或者最近五日内最高, 卖出
        if (data["Close"][i] > data["Close"][i - 1] and (data["Close"][i] - data["Close"][i - 1]) / data["Close"][
            i - 1] >= buy_rate) or \
                (data["Close"][i] == max(data["Close"][i - 5:i + 1])):
            return Sell
        # 相对昨日的下跌率超过5%, 或者最近五日内最低, 买入
        if (data["Close"][i] < data["Close"][i - 1] and (data["Close"][i - 1] - data["Close"][i]) / data["Close"][
            i - 1] >= sell_rate) or \
                (data["Close"][i] == min(data["Close"][i - 5:i + 1])):
            return Buy
    return Wait


class Investor:
    def __init__(self, No, method, adoptCoff, volumeCoff, cap, startTime, stockCode):
        self.No = No
        self.strategy = method  # 投资采取策略
        self.adoptCoff = adoptCoff  # 采纳技术指标信号的可能性
        self.volumeCoff = volumeCoff  # 交易量系数
        self.initialCapital = cap  # 初始资产
        self.cash = cap  # 持有现金额度(单位, 元)
        self.capital = cap  # 总资产(单位, 元)
        self.stockHolding = [0 for _ in range(300)]  # 股票持有数量(单位, 股)
        self.frequencyBuy = [0 for _ in range(300)]  # 买入频次
        self.frequencySell = [0 for _ in range(300)]  # 卖出频次
        self.dateTime = startTime
        self.log = []
        self.infoIndex = 0
        self.Map = {
            "MACD": onMACD, "DMI": onDMI,
            "PSY": onPSY, "VR": onVR,
            "KDJ": onKDJ, "WR": onWR,
            "RSI": onRSI, "BIAS": onBIAS,
            "ARBR": onARBR, "BOLL": onBOLL,
            "ROC": onROC, "MTM": onMTM,
            "NULL": onNULL, "CUKD": CUKD,
            "HTLS": HTLS,
        }
        self.stockCode = stockCode
        self.fee = [0 for _ in range(300)]  # 交易过程中产生的费用
        self.successBuy = [0 for _ in range(300)]  # 成功买入次数
        self.successSell = [0 for _ in range(300)]  # 成功卖出次数
        self.lastBuyPrice = [0 for _ in range(300)]     # 上次买入价格
        self.lastSellPrice = [0 for _ in range(300)]    # 上次卖出价格
        self.cumulativeBuy = [0 for _ in range(300)]    # 累计买入股票数
        self.cumulativeSell = [0 for _ in range(300)]   # 累计卖出股票数
        self.stopOutFee = 0         # 最终强制平仓后的总费用
        self.stopOutCapital = 0     # 最终强制平仓后的总资产

    # 判断是否要交易(i表示股票序号)
    def analyse(self, stockIndex):
        if int(stockInfo[stockIndex].data["Date"][self.dateTime[stockIndex]][0:9]) > int(stockInfo[0].data["Date"][self.dateTime[0]][0:9]):
            return
        methods = self.strategy.split("-")
        signals = []
        for i in range(0, len(methods)):
            signals.append(self.Map[methods[i]](self.dateTime[stockIndex], stockIndex))
        # 各技术指标信号相同, 则采取此动作
        if len(set(signals)) == 1:
            self.deal(signals[0], stockIndex)
        else:
            self.deal(Wait, stockIndex)

    """
        服务费:
        (1) 印花税: 0.1% - 0.001
        (2) 过户费: 0.002% - 0.00002
        (3) 交易佣金: 0.03% - 0.0003 (5起步)
        (4) 证管费: 0.002% - 0.00002
        (5) 经手费: 0.00487% - 0.0000487
    """

    # 交易
    def deal(self, signal, stockIndex):
        data = stockInfo[stockIndex].data
        # 因技术指标因素买入
        if self.cash > 0 and signal == Buy:
            ratio = self.amount(1, stockIndex) / 100
            dealAmount = math.floor(self.cash * ratio / data["Close"][self.dateTime[stockIndex]] / 100) * 100
            while (1 + 0.00002 + 0.00002 + 0.0000487) * dealAmount * data["Close"][self.dateTime[stockIndex]] + max(5, 0.0003 * dealAmount * data["Close"][self.dateTime[stockIndex]]) > self.cash:
                dealAmount -= 100
                if dealAmount <= 0:
                    return
            if dealAmount <= 0:
                return
            self.lastBuyPrice[stockIndex] = data["Close"][self.dateTime[stockIndex]]
            # 本次买入价格低于最近卖出价格, 记为成功交易
            if data["Close"][self.dateTime[stockIndex]] < self.lastSellPrice[stockIndex]:
                self.successBuy[stockIndex] += 1
            self.cumulativeBuy[stockIndex] += dealAmount
            self.stockHolding[stockIndex] += dealAmount
            self.cash -= (1 + 0.00002 + 0.00002 + 0.0000487) * dealAmount * data["Close"][self.dateTime[stockIndex]] + max(5, 0.0003 * dealAmount * data["Close"][self.dateTime[stockIndex]])
            self.fee[stockIndex] += max(5, 0.0003 * dealAmount * data["Close"][self.dateTime[stockIndex]]) + (0.00002 + 0.00002 + 0.0000487) * dealAmount * data["Close"][self.dateTime[stockIndex]]
            self.frequencyBuy[stockIndex] += 1
        # 因技术指标卖出
        elif self.stockHolding[stockIndex] > 0 and signal == Sell:
            ratio = self.amount(2, stockIndex) / 100
            dealAmount = math.floor(self.stockHolding[stockIndex] * ratio / 100) * 100
            while self.cash + (1 - 0.001 - 0.00002 - 0.00002 - 0.0000487) * dealAmount * data["Close"][self.dateTime[stockIndex]] - max(5, 0.0003 * dealAmount * data["Close"][self.dateTime[stockIndex]]) < 0:
                dealAmount -= 100
                if dealAmount <= 0:
                    return
            if dealAmount <= 0:
                return
            # 更新最近卖出的价格
            self.lastSellPrice[stockIndex] = data["Close"][self.dateTime[stockIndex]]
            # 判断是否是成功交易
            if data["Close"][self.dateTime[stockIndex]] > self.lastBuyPrice[stockIndex]:  # 本次买入价格低于最近卖出价格, 记为成功交易
                self.successSell[stockIndex] += 1
            # 该类股票累计买入数量
            self.cumulativeSell[stockIndex] += dealAmount
            # 持有股票数减少
            self.stockHolding[stockIndex] -= dealAmount
            # 持有现金增加
            self.cash += (1 - 0.001 - 0.00002 - 0.00002 - 0.0000487) * dealAmount * data["Close"][self.dateTime[stockIndex]] - max(5, 0.0003 * dealAmount * data["Close"][self.dateTime[stockIndex]])
            # 记录交易费用
            self.fee[stockIndex] += max(5, 0.0003 * dealAmount * data["Close"][self.dateTime[stockIndex]]) + (0.001 + 0.00002 + 0.00002 + 0.0000487) * dealAmount * data["Close"][self.dateTime[stockIndex]]
            self.frequencySell[stockIndex] += 1

    # 决定交易量, tpe指示交易类型, 1-依据指标买入系数进行交易, 2-依据指标卖出系数进行交易, 3-根据非指标买入系数进行交易, 4-根据非指标卖出系数进行交易
    def amount(self, tpe, stockIndex):
        happen = random.random()
        percentage = 0
        if happen <= self.adoptCoff:
            percentage = random.random() * 10
            while True:
                if percentage >= 100:
                    percentage = 100
                    break
                if random.random() <= self.volumeCoff:
                    percentage += random.random() * 10
                else:
                    break
        return percentage

    # 结算盈亏
    def settlement(self):
        for stockIndex in range(0, len(self.stockCode)):
            # 如果还没有开始允许投资, 则跳过
            if stockIndex != 0 and int(stockInfo[stockIndex].data["Date"][self.dateTime[stockIndex]][0:9]) >= int(stockInfo[0].data["Date"][self.dateTime[0]][0:9]):
                # eg: 股票A的起始日期是8.6, 1号股票后移一天才8.6, 这时不能给A股票加1天
                continue
            self.dateTime[stockIndex] += 1

        self.capital = self.cash
        for stockIndex in range(0, len(self.stockCode)):
            if int(stockInfo[stockIndex].data["Date"][self.dateTime[stockIndex]][0:9]) > int(stockInfo[0].data["Date"][self.dateTime[0]][0:9]):
                continue
            data = stockInfo[stockIndex].data
            self.capital += self.stockHolding[stockIndex] * data["Close"][self.dateTime[stockIndex]]


def simulate():
    print("数据准备 ...")
    """ 准备股票数据及指标数据 """
    global stockInfo
    workbook = xlrd.open_workbook("../data/All_Stock_Info(hushen300).xls")
    stockCode = workbook.sheet_names()
    print("准备data ...")
    for i in range(0, len(stockCode)):
        print(str(i+1) + "/" + str(len(stockCode)))
        gi.generateIndex(workbook, stockCode[i])
        stockInfo.append(copy.deepcopy(gi.stock))

    """ 每支股票的起始行数 """
    print("计算起始行数 ...")
    startRow = []
    for i in range(0, len(stockCode)):
        print(str(i + 1) + "/" + str(len(stockCode)))
        startRow.append(first_Row(workbook, stockCode[i]) - 1)

    """ 股票投资策略 """
    methods = []
    basic_methods = ["MACD", "DMI", "PSY", "VR", "KDJ", "WR", "RSI", "BIAS", "ARBR", "BOLL", "ROC", "MTM"]
    for i in range(0, len(basic_methods)):
        methods.append(basic_methods[i])
    for i in range(0, len(basic_methods)):
        for j in range(i + 1, len(basic_methods)):
            if i != j:
                methods.append(basic_methods[i] + "-" + basic_methods[j])
    methods.insert(0, "NULL")
    """ 准备各投资者数据 """
    capital = [random.random() for _ in range(10000)]
    adoptCoff = [random.random() for _ in range(10000)]
    volumeCoff = [random.random() for _ in range(10000)]
    for j in range(0, 9700):
        capital[j] = capital[j] * 500000
    for j in range(9700, 10000):
        capital[j] = 500000 + capital[j] * 10000000

    """ 对各投资策略进行投资模拟 """
    print("模拟投资 ...")
    for i in range(0, len(methods)):
        method = methods[i]
        res = []
        for k in range(0, 1000):
            print(str(i+1) + "/" + str(len(methods)) + "  " + method + "  " + str(k + 1) + "/" + "10000")
            # def __init__(self, No, method, adoptCoff, volumeCoff, cap, startTime, stockCode):
            example = Investor(k, method, adoptCoff[k], volumeCoff[k], capital[k], copy.deepcopy(startRow), stockCode)
            for t in range(startRow[0], len(stockInfo[0].data["Date"]) - 1):  # 考查每一个交易日
                for stockIndex in range(0, len(stockCode)):  # 在交易日t, 对每支股票分别决策
                    example.analyse(stockIndex)
                example.settlement()
            # 强制平仓
            example.stopOutCapital = example.capital    # 强制平仓后的总资金
            example.stopOutFee = sum(example.fee)  # 强制平仓后缴纳的总费用
            for stockIndex in range(0, len(stockCode)):
                if example.stockHolding[stockIndex] != 0:
                    example.stopOutCapital -= (0.001 + 0.00002 + 0.00002 + 0.0000487) * example.stockHolding[stockIndex] * stockInfo[stockIndex].data["Close"][-1] + max(5, 0.0003 * example.stockHolding[stockIndex] * stockInfo[stockIndex].data["Close"][-1])
                    example.stopOutFee += (0.001 + 0.00002 + 0.00002 + 0.0000487) * example.stockHolding[stockIndex] * stockInfo[stockIndex].data["Close"][-1] + max(5, 0.0003 * example.stockHolding[stockIndex] * stockInfo[stockIndex].data["Close"][-1])
            info = {"No": example.No,
                    "Initial_Capital": example.initialCapital,
                    "Capital": example.capital,
                    "StopOutCapital": example.stopOutCapital,
                    "Annual_Profit_Rate": (example.capital - example.initialCapital) / example.initialCapital / 3,
                    "StopOut_Annual_Profit_Rate": (example.stopOutCapital - example.initialCapital) / example.initialCapital / 3,
                    "Cash": example.cash,
                    "AdoptCoff": example.adoptCoff,
                    "VolumeCoff": example.volumeCoff,
                    "StockHolding": sum(example.stockHolding),
                    "FrequencyBuy": sum(example.frequencyBuy),
                    "FrequencySell": sum(example.frequencySell),
                    "Frequency": sum(example.frequencyBuy) + sum(example.frequencySell),
                    "Fee": sum(example.fee),
                    "StopOutFee": example.stopOutFee,
                    "SuccessBuy": sum(example.successBuy),
                    "SuccessSell": sum(example.successSell),
                    "Success": sum(example.successBuy) + sum(example.successSell),
                    "SuccessRate": (sum(example.successBuy) + sum(example.successSell)) / (sum(example.frequencyBuy) + sum(example.frequencySell) + 0.001)}
            for stockIndex in range(0, len(stockCode)):
                info["StockHolding" + stockCode[stockIndex]] = example.stockHolding[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                info["CumulativeBuy" + stockCode[stockIndex]] = example.cumulativeBuy[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                info["CumulativeSell" + stockCode[stockIndex]] = example.cumulativeSell[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                info["FrequencyBuy" + stockCode[stockIndex]] = example.frequencyBuy[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                info["FrequencySell" + stockCode[stockIndex]] = example.frequencySell[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                info["Fee" + stockCode[stockIndex]] = example.fee[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                if example.stockHolding[stockIndex] != 0:
                    info["StopOutFee" + stockCode[stockIndex]] = example.fee[stockIndex] + (0.001 + 0.00002 + 0.00002 + 0.0000487) * example.stockHolding[stockIndex] * stockInfo[stockIndex].data["Close"][-1] + max(5, 0.0003 * example.stockHolding[stockIndex] * stockInfo[stockIndex].data["Close"][-1])
                else:
                    info["StopOutFee" + stockCode[stockIndex]] = example.fee[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                info["StopOutFrequencyBuy" + stockCode[stockIndex]] = example.frequencyBuy[stockIndex]
            for stockIndex in range(0, len(stockCode)):
                if example.stockHolding[stockIndex] != 0:
                    info["StopOutFrequencySell" + stockCode[stockIndex]] = example.frequencySell[stockIndex] + 1
                else:
                    info["StopOutFrequencySell" + stockCode[stockIndex]] = example.frequencySell[stockIndex]
            res.append(info)
        df = pd.DataFrame(res)
        df.to_excel("../dataSimulation/" + method + ".xlsx")


if __name__ == "__main__":
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
    plt.rcParams["axes.unicode_minus"] = False
    simulate()
