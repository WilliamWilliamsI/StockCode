import math
import xlrd
import xlwt
import pandas as pd
import numpy as np


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

            # "OBV_OBV": [],
            # "OBV_MAOBV": [],

            "SAR_Ascending_SAR": [],
            "SAR_Descending_SAR": [],
            "SAR_AF": [],
            "SAR_SAR": [],  # 最终选取的SAR
            "SAR_Feature": [],  # 涨跌状态

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


stock = Stock()

"""数据校正: https://finance.sina.com.cn/realstock/company/sh000300/nc.shtml
沪深300指数"""


def variance(arr: list):
    """
    * 计算方差
    * :param
    *   arr: 数组
    * :return数组方差
    """
    sum = 0
    res = 0
    for d in arr:
        sum += d
    mean = sum / len(arr)
    for d in arr:
        res += (d - mean) ** 2
    return res / (len(arr) - 1)


def conditionPush(arr: list, i, ele):
    """
    * 条件插入数据(鲁棒性，防止重复插入)
    * :param
    *   arr: 数组
    *   i: 更新位置
    * :return NULL
    """
    if i >= len(arr):
        arr.append(ele)
    else:
        arr[i] = ele


def conditionGet(arr: list, ele: str):
    """
    * 条件读取数据(鲁棒性，防止float('')报错)
    * :param
    *   arr: 数组
    *   ele: 字符
    * :return NULL
    """
    if ele != '':
        arr.append(float(ele))
    else:
        arr.append(0)


def MA(i, N):
    """
    * N日移动平均线: N日移动平均线 = N日收市价之和/N
    * :param
    *   i: 序列号(当天所处天数)
    *   N: 天数
    * :return 移动平均线值
    """
    data = stock.data
    sum = 0
    dataSlice = data["Close"][i - N + 1:i + 1]
    for d in dataSlice:
        sum += d
    # data length is less than N
    if i < N - 1:
        return sum / (i - 1)
    else:
        return sum / N


def MACD(i, SHORT=12, LONG=26, M=9):
    """ double Checked
    * N日MACD(指数平滑移动平均线), 数据更新到data
    * reference: https://www.jianshu.com/p/24f992e457ca
    * https://blog.csdn.net/u014496893/article/details/122776058
    * https://zhuanlan.zhihu.com/p/361132689
    * 代码:http://www.cppblog.com/gaimor/archive/2016/08/30/214240.html
    * :param
    *   i: 序列号(当天所处天数)
    *   N: 天数
    * :return 当天MACD值(bar值大小)
    """
    data = stock.data
    DIF, DEA, MACD, EMA12, EMA26 = 0, 0, 0, 0, 0
    if i >= 1:
        EMA12 = data["MACD_EMA12"][i - 1] * (SHORT - 1) / (SHORT + 1) + data["Close"][i] * 2 / (SHORT + 1)
        EMA26 = data["MACD_EMA26"][i - 1] * (LONG - 1) / (LONG + 1) + data["Close"][i] * 2 / (LONG + 1)
        DIF = EMA12 - EMA26
        DEA = data["MACD_DEA"][i - 1] * (M - 1) / (M + 1) + DIF * 2 / (M + 1)
        MACD = 2 * (DIF - DEA)
        # DIF = (data["Close"][i] * 2 / (SHORT+1) + data["MACD_DIF"][i-1] * (SHORT-1) / (SHORT+1)) - (
        #         data["Close"][i] * 2 / (LONG+1) + data["MACD_DIF"][i-1] * (LONG-1) / (LONG+1))
        # conditionPush(data["MACD_DIF"], i, DIF)
        # DEA = data["MACD_DIF"][i] * 2 / (M+1) + data["MACD_DEA"][i-1] * (M-1) / (M+1)
        # conditionPush(data["MACD_DEA"], i, DEA)
        # MACD = 2 * (DIF - DEA)
        conditionPush(data["MACD_EMA12"], i, EMA12)
        conditionPush(data["MACD_EMA26"], i, EMA26)
        conditionPush(data["MACD_DIF"], i, DIF)
        conditionPush(data["MACD_DEA"], i, DEA)
        conditionPush(data["MACD_MACD"], i, MACD)
    else:
        conditionPush(data["MACD_EMA12"], i, EMA12)
        conditionPush(data["MACD_EMA26"], i, EMA26)
        conditionPush(data["MACD_DIF"], i, DIF)
        conditionPush(data["MACD_DEA"], i, DEA)
        conditionPush(data["MACD_MACD"], i, MACD)
    return [MACD, DIF, DEA]


def DMI(i, M1=14, M2=6):
    """ double checked
    * DMI动向指标(趋向指标)
    * reference:
    * https://stock.gucheng.com/202108/4092419.shtml
    * https://baike.baidu.com/item/DMI%E6%8C%87%E6%A0%87/3423254?fr=aladdin
    * source code:
        TR=SUM(MAX(MAX(HIGH-LOW,ABS(HIGH-REF(CLOSE,1))),ABS(LOW-REF(CLOSE,1))),M1);
        HD=HIGH-REF(HIGH,1);
        LD=REF(LOW,1)-LOW;
        DMP=SUM(IF(HD>0 AND HD>LD,HD,0),M1);
        DMM=SUM(IF(LD>0 AND LD>HD,LD,0),M1);
        DI1:DMP*100/TR;
        DI2:DMM*100/TR;
        ADX:MA(ABS(DI2-DI1)/(DI1+DI2)*100,M2);
        ADXR:(ADX+REF(ADX,M2))/2;
    * :param
    * :return
    """
    data = stock.data
    TR = 0
    if i - M1 + 1 >= 0:
        for ii in range(i - M1 + 1, i + 1):
            tmp = max(data["High"][ii] - data["Low"][ii], abs(data["High"][ii] - data["Close"][ii - 1]))
            TR += max(tmp, abs(data["Low"][ii] - data["Close"][ii - 1]))
    conditionPush(data["DMI_TR"], i, TR)
    DMP, DMM = 0, 0
    if i - M1 + 1 >= 0:
        for ii in range(i - M1 + 1, i + 1):
            HD = data["High"][ii] - data["High"][ii - 1]
            LD = data["Low"][ii - 1] - data["Low"][ii]
            if HD > 0 and HD > LD:
                DMP += HD
            if LD > 0 and LD > HD:
                DMM += LD
    DI1, DI2 = 0, 0
    if TR != 0:
        DI1 = DMP * 100 / TR
        DI2 = DMM * 100 / TR
    conditionPush(data["DMI_DI1"], i, DI1)
    conditionPush(data["DMI_DI2"], i, DI2)
    ADX = 0
    # 使用2*M1是为了保证向前取的data[DI1]和data[DI2]均不为0
    if i - 2 * M1 + 1 >= 0:
        for ii in range(i - M2 + 1, i + 1):
            ADX += abs(data["DMI_DI2"][ii] - data["DMI_DI1"][ii]) / (data["DMI_DI1"][ii] + data["DMI_DI2"][ii]) * 100
        ADX /= M2
    conditionPush(data["DMI_ADX"], i, ADX)
    ADXR = 0
    if i - 2 * M1 + 1 >= 0:
        ADXR = (data["DMI_ADX"][i] + data["DMI_ADX"][i - M2]) / 2
    conditionPush(data["DMI_ADXR"], i, ADXR)


def EXPMA(i, P1=5, P2=10, P3=20, P4=60):
    """
    * :param
    *   i: index
    *   P1: 5
    *   P2: 10
    *   P3: 20
    *   P4: 60
    * :return
    *   [MA1, MA2, MA3, MA4]
    * code:
    *   EMA: 求指数平滑移动平均, Y = [2*X+(N-1)*Y']/(N+1)
    *   MA1:EMA(CLOSE,P1);
    *   MA2:EMA(CLOSE,P2);
    *   MA3:EMA(CLOSE,P3);
    *   MA4:EMA(CLOSE,P4);
    *
    """
    data = stock.data
    MA1, MA2, MA3, MA4 = 0, 0, 0, 0
    if i - 1 >= 0:
        MA1 = 2 * data["Close"][i] / (P1 + 1) + (P1 - 1) * data["EXPMA_MA1"][i - 1] / (P1 + 1)
        MA2 = 2 * data["Close"][i] / (P2 + 1) + (P2 - 1) * data["EXPMA_MA2"][i - 1] / (P2 + 1)
        MA3 = 2 * data["Close"][i] / (P3 + 1) + (P3 - 1) * data["EXPMA_MA3"][i - 1] / (P3 + 1)
        MA4 = 2 * data["Close"][i] / (P4 + 1) + (P4 - 1) * data["EXPMA_MA4"][i - 1] / (P4 + 1)
    conditionPush(data["EXPMA_MA1"], i, MA1)
    conditionPush(data["EXPMA_MA2"], i, MA2)
    conditionPush(data["EXPMA_MA3"], i, MA3)
    conditionPush(data["EXPMA_MA4"], i, MA4)
    return [MA1, MA2, MA3, MA4]


def PSY(i, N=12, M=6):
    """
    * PSY: 心理线
    * :param
    * :return
    * code: PSY:COUNT(CLOSE>REF(CLOSE,1),N)/N*100;
    """
    data = stock.data
    PSY, PSYMA, count = 0, 0, 0
    if i - N >= 0:
        for ii in range(i - N + 1, i + 1):
            if data["Close"][ii] > data["Close"][ii - 1]:
                count += 1
        PSY = count / N * 100
        PSYMA = data["PSY_PSYMA"][i - 1] * (M - 1) / (M + 1) + PSY * 2 / (M + 1)
    conditionPush(data["PSY_PSY"], i, PSY)
    conditionPush(data["PSY_PSYMA"], i, PSYMA)
    return [PSY, PSYMA]


def KDJ(i, N=9):
    """ Checked
    * KDJ
    * reference: https://baike.baidu.com/item/KDJ%E6%8C%87%E6%A0%87/6328421?fromtitle=kdj&fromid=3423560&fr=aladdin
    * :param
    *   i: 索引号
    *   N: KDJ天数
    """
    data = stock.data
    if i < N - 1:
        K = 50
        D = 50
        J = 3 * K - 2 * D
        conditionPush(data["KDJ_K"], i, K)
        conditionPush(data["KDJ_D"], i, D)
        conditionPush(data["KDJ_J"], i, J)
    else:
        Ln = min(data["Low"][i - N + 1: i + 1])
        Hn = max(data["High"][i - N + 1: i + 1])
        RSV = (data["Close"][i] - Ln) / (Hn - Ln) * 100
        K = 2 / 3 * data["KDJ_K"][i - 1] + 1 / 3 * RSV
        D = 2 / 3 * data["KDJ_D"][i - 1] + 1 / 3 * K
        J = 3 * K - 2 * D
        conditionPush(data["KDJ_K"], i, K)
        conditionPush(data["KDJ_D"], i, D)
        conditionPush(data["KDJ_J"], i, J)
        return [K, D, J]


def WR(i, N1=6, N2=10):
    """ double Checked
    * 威廉超买超卖指标WR
    * reference: https://stock.gucheng.com/202109/4098096.shtml
    * :param
    *   i: 序列号
    *   N: 周期参数(6, 10 recommended)
    * :return [WR1, WR2]
    """
    data = stock.data
    if i < N1 - 1 or i < N2 - 1:
        WR1 = 0
        WR2 = 0
        conditionPush(data["WR1"], i, WR1)
        conditionPush(data["WR2"], i, WR2)
        return [WR1, WR2]
    else:
        Hn1 = max(data["High"][i - N1 + 1: i + 1])
        Ln1 = min(data["Low"][i - N1 + 1: i + 1])
        Hn2 = max(data["High"][i - N2 + 1: i + 1])
        Ln2 = min(data["Low"][i - N2 + 1: i + 1])
        WR1 = 100 * (Hn1 - data["Close"][i]) / (Hn1 - Ln1)
        WR2 = 100 * (Hn2 - data["Close"][i]) / (Hn2 - Ln2)
        conditionPush(data["WR1"], i, WR1)
        conditionPush(data["WR2"], i, WR2)
        return [WR1, WR2]


def RSI(i, N1=6, N2=12, N3=24):
    """
    * RSI: 相对强弱指标
    * :param
    *   N1: 6
    *   N2: 12
    *   N3: 24
    * :return
    * code:
    * LC := REF(CLOSE,1);
    * RSI$1:SMA(MAX(CLOSE-LC,0),N1,1)/SMA(ABS(CLOSE-LC),N1,1)*100;
    注: SMA(MAX(CLOSE-LC,0),N1,1)是指对此变量进行移动平均，并非对RSI$1整个进行移动平均
    * RSI$2:SMA(MAX(CLOSE-LC,0),N2,1)/SMA(ABS(CLOSE-LC),N2,1)*100;
    * RSI$3:SMA(MAX(CLOSE-LC,0),N3,1)/SMA(ABS(CLOSE-LC),N3,1)*100;
    * a:20;
    * d:80;
    """
    data = stock.data
    RSI6, RSI12, RSI24, tmp1_1, tmp1_2, tmp2_1, tmp2_2, tmp3_1, tmp3_2 = 0, 0, 0, 0, 0, 0, 0, 0, 0
    if i - N3 + 1 >= 0:
        LC = data["Close"][i - 1]
        tmp1_1 = max(data["Close"][i] - LC, 0) * 1 / N1 + data["RSI_tmp1_1"][i - 1] * (N1 - 1) / N1
        tmp1_2 = abs(LC - data["Close"][i]) * 1 / N1 + data["RSI_tmp1_2"][i - 1] * (N1 - 1) / N1
        if tmp1_2 == 0:
            print("division by 0 in RSI tmp1_2")
            tmp1_2 = 0.000001
        RSI6 = tmp1_1 / tmp1_2 * 100
        tmp2_1 = max(data["Close"][i] - LC, 0) * 1 / N2 + data["RSI_tmp2_1"][i - 1] * (N2 - 1) / N2
        tmp2_2 = abs(LC - data["Close"][i]) * 1 / N2 + data["RSI_tmp2_2"][i - 1] * (N2 - 1) / N2
        if tmp2_2 == 0:
            print("division by 0 in RSI tmp2_2")
            tmp2_2 = 0.000001
        RSI12 = tmp2_1 / tmp2_2 * 100
        tmp3_1 = max(data["Close"][i] - LC, 0) * 1 / N3 + data["RSI_tmp3_1"][i - 1] * (N3 - 1) / N3
        tmp3_2 = abs(LC - data["Close"][i]) * 1 / N3 + data["RSI_tmp3_2"][i - 1] * (N3 - 1) / N3
        if tmp3_2 == 0:
            print("division by 0 in RSI tmp3_2")
            tmp3_2 = 0.000001
        RSI24 = tmp3_1 / tmp3_2 * 100
    conditionPush(data["RSI_tmp1_1"], i, tmp1_1)
    conditionPush(data["RSI_tmp1_2"], i, tmp1_2)
    conditionPush(data["RSI_tmp2_1"], i, tmp2_1)
    conditionPush(data["RSI_tmp2_2"], i, tmp2_2)
    conditionPush(data["RSI_tmp3_1"], i, tmp3_1)
    conditionPush(data["RSI_tmp3_2"], i, tmp3_2)
    conditionPush(data["RSI_RSI6"], i, RSI6)
    conditionPush(data["RSI_RSI12"], i, RSI12)
    conditionPush(data["RSI_RSI24"], i, RSI24)
    return [RSI6, RSI12, RSI24]


def BIAS(i, N1=6, N2=12, N3=24):
    """
    * 乖离率bias
    * reference:https://baike.baidu.com/item/%E4%B9%96%E7%A6%BB%E7%8E%87/420286?fr=aladdin
    * :param
    *   i: 索引号
    *   N: 周期(6, 12, 24 recommended)
    """
    data = stock.data
    BIAS1, BIAS2, BIAS3 = 0, 0, 0
    if i >= N3 - 1:
        MA1 = MA(i, N1)
        MA2 = MA(i, N2)
        MA3 = MA(i, N3)
        BIAS1 = (data["Close"][i] - MA1) / MA1 * 100
        BIAS2 = (data["Close"][i] - MA2) / MA2 * 100
        BIAS3 = (data["Close"][i] - MA3) / MA3 * 100
    conditionPush(data["BIAS_BIAS1"], i, BIAS1)
    conditionPush(data["BIAS_BIAS2"], i, BIAS2)
    conditionPush(data["BIAS_BIAS3"], i, BIAS3)
    return [BIAS1, BIAS2, BIAS3]


def VR(i, M1=26, M2=100, M3=200):
    """
    * VR容量比率
    * :param
    * i: index
    * M1: 26
    * M2: 100
    * M3: 200
    * :return
    * [VR, A=M2, B=M3]
    * code
    * TH:=SUM(IF(CLOSE>REF(CLOSE,1),VOL,0),M1);
    * TL:=SUM(IF(CLOSE<REF(CLOSE,1),VOL,0),M1);
    * TQ:=SUM(IF(CLOSE=REF(CLOSE,1),VOL,0),M1);
    * VR:100*(TH*2+TQ)/(TL*2+TQ);
    """
    data = stock.data
    TH, TL, TQ, VR = 0.01, 0.01, 0.01, 0.01
    if i - M1 + 1 >= 0:
        for ii in range(i - M1 + 1, i + 1):
            if data["Close"][ii] > data["Close"][ii - 1]:
                TH += data["Volume"][ii]
            elif data["Close"][ii] < data["Close"][ii - 1]:
                TL += data["Volume"][ii]
            else:
                TQ += data["Volume"][ii]
    VR = 100 * (TH * 2 + TQ) / (TL * 2 + TQ)
    conditionPush(data["VR_VR"], i, VR)
    return [VR, M2, M3]


def ARBR(i, M1=26, M2=70, M3=150):
    """
    * ABBR: 人气意愿指标
    * :param
    *   i: index
    *   M1: 26
    *   M2: 70
    *   M3: 150
    * :return
    *   [AR, BR, A=M2, B=M3]
    * code:
    *    AR:SUM(HIGH-OPEN,M1)/SUM(OPEN-LOW,M1)*100;
    *    BR:SUM(MAX(0,HIGH-REF(CLOSE,1)),M1)/SUM(MAX(0,REF(CLOSE,1)-LOW),M1)*100;
    *    a:M2;
    *    b:M3;
    """
    data = stock.data
    AR, BR = 0, 0
    SUM1, SUM2, SUM3, SUM4 = 0.0001, 0.0001, 0.0001, 0.0001
    if i - M1 + 1 >= 0:
        for ii in range(i - M1 + 1, i + 1):
            SUM1 += data["High"][ii] - data["Open"][ii]
            SUM2 += data["Open"][ii] - data["Low"][ii]
        AR = SUM1 / SUM2 * 100
        for ii in range(i - M1 + 1, i + 1):
            SUM3 += max(0, data["High"][ii] - data["Close"][ii - 1])
            SUM4 += max(0, data["Close"][ii - 1] - data["Low"][ii])
        BR = SUM3 / SUM4 * 100
    conditionPush(data["ARBR_AR"], i, AR)
    conditionPush(data["ARBR_BR"], i, BR)
    return [AR, BR, M2, M3]


def ASI(i, M1=26, M2=10):
    """
    * ASI: 振动升降指标
    * :param
    *   i: index
    *   M1: 26
    *   M2: 10
    * :return
    *   [ASI, ASIT]
    * code:
    *     LC=REF(CLOSE,1);
    *     AA=ABS(HIGH-LC);
    *     BB=ABS(LOW-LC);
    *     CC=ABS(HIGH-REF(LOW,1));
    *     DD=ABS(LC-REF(OPEN,1));
    *     R=IF(AA>BB AND AA>CC,AA+BB/2+DD/4,IF(BB>CC AND BB>AA,BB+AA/2+DD/4,CC+DD/4));
    *     X=(CLOSE-LC+(CLOSE-OPEN)/2+LC-REF(OPEN,1));
    *     SI=50*X/R*MAX(AA,BB)/3;
    *     ASI:SUM(SI,M1);
    *     ASIT:MA(ASI,M2);
    """
    data = stock.data
    SI, ASI, ASIT, tmp = 0, 0, 0, 0
    if i - 1 >= 0:
        LC = data["Close"][i - 1]
        AA = abs(data["High"][i] - LC)
        BB = abs(data["Low"][i] - LC)
        CC = abs(data["High"][i] - data["Low"][i - 1])
        DD = abs(LC - data["Open"][i - 1])
        R = 0
        if AA > BB and AA > CC:
            R = AA + BB / 2 + DD / 4
        if BB > CC and BB > AA:
            R = BB + AA / 2 + DD / 4
        else:
            R = CC + DD / 4
        if R == 0:
            R = 0.00001
        X = data["Close"][i] - LC + (data["Close"][i] - data["Open"][i]) / 2 + LC - data["Open"][i - 1]
        SI = 50 * X / R * max(AA, BB) / 3
    conditionPush(data["ASI_SI"], i, SI)

    if i - M1 + 1 >= 0:
        for ii in range(i - M1 + 1, i + 1):
            ASI += data["ASI_SI"][ii]
    conditionPush(data["ASI_ASI"], i, ASI)

    if i - M2 + 1 >= 0:
        for ii in range(i - M2 + 1, i + 1):
            tmp += data["ASI_ASI"][ii]
        ASIT = tmp / M2
    conditionPush(data["ASI_ASIT"], i, ASIT)


""" ???????
* OBV: 能量潮
* :param 
*   i: index
*   M1: 30
* :return
*   [OBV, MAOBV]
* code:
*   OBV:SUM(IF(CLOSE>REF(CLOSE,1),VOL,IF(CLOSE<REF(CLOSE,1),-VOL,0)),0)/10000;
*   MAOBV:MA(OBV,M1)
* 
"""

# todo 修正
# def OBV(i, M1=30):
#     SUM, OBV, MAOBV = 0, 0, 0
#     for ii in range(1, i+1):
#         if data["Close"][ii] > data["Close"][ii-1]:
#             SUM += data["Volume"][ii]
#         elif data["Close"][ii] < data["Close"][ii-1]:
#             SUM -= data["Volume"][ii]
#     OBV = SUM / 10000
#     conditionPush(data["OBV_OBV"], i, OBV)
#     if i-M1+1 >= 0:
#         for ii in range(i-M1+1, i+1):
#             MAOBV += data["OBV_OBV"][ii]
#         MAOBV /= M1
#     conditionPush(data["OBV_MAOBV"], i, MAOBV)
#     return [OBV, MAOBV]


"""SAR: 抛物转向
* :param
*   i: index
*   N: 4
* :return
*
* code:
*   SAR_S(n,2,20)
* reference: https://www.zhihu.com/question/65785061
"""

# def SAR(i, N=4):
#     H1 = data["High"][i]
#     L1 = data["Low"][i]
#     if i == 0:
#         Ascending_SAR1 = data["Low"][i]
#         Descending_SAR1 = data["High"][i]
#         AF = 0.02
#         Ascending_SAR2 = Ascending_SAR1 + AF * (H1 - Ascending_SAR1)
#         Descending_SAR2 = Descending_SAR1 + AF * (L1 - Descending_SAR1)
#         conditionPush(data["SAR_Ascending_SAR"], i, Ascending_SAR2)
#         conditionPush(data["SAR_Descending_SAR"], i, Descending_SAR2)
#         conditionPush(data["SAR_SAR"], i, data["Close"][i])
#         conditionPush(data["SAR_AF"], i, AF)
#         conditionPush(data["SAR_Feature"], i, "ascend")
#
#     else:
#         Ascending_SAR2 = data["SAR_Ascending_SAR"][i-1] + data["SAR_AF"][i-1] * (H1 - data["SAR_Ascending_SAR"][i-1])
#         Descending_SAR2 = data["SAR_Descending_SAR"][i-1] + data["SAR_AF"][i-1] * (L1 - data["SAR_Descending_SAR"][i-1])
#         conditionPush(data["SAR_Ascending_SAR"], i, Ascending_SAR2)
#         conditionPush(data["SAR_Descending_SAR"], i, Descending_SAR2)
#         if H1 > data["SAR_Ascending_SAR"][i-1]:
#             conditionPush(data["SAR_SAR"], i, Ascending_SAR2)
#             if data["SAR_Feature"][i-1] == "ascend" and H1 > max(data["High"][:i]):
#                 AF = min(data["SAR_AF"][i-1] * 2, 0.2)
#                 conditionPush(data["SAR_AF"], i, AF)
#                 conditionPush(data["SAR_Feature"], i, "ascend")
#             elif data["SAR_Feature"][i-1] == "descend" and H1 > max(data["High"][:i]):
#                 AF = 0.02
#                 conditionPush(data["SAR_AF"], i, AF)
#                 conditionPush(data["SAR_Feature"], i, "ascend")
#         elif L1 < data["SAR_Descending_SAR"][i-1]:
#             conditionPush(data["SAR_SAR"], i, Descending_SAR2)
#             if data["SAR_Feature"][i-1] == "descend" and L1 < min(data["Low"][:i]):
#                 AF = min(data["SAR_AF"][i-1] * 2, 0.2)
#                 conditionPush(data["SAR_AF"], i, AF)
#                 conditionPush(data["SAR_Feature"], i, "descend")
#             elif data["SAR_Feature"][i-1] == "ascend" and L1 < min(data["Low"][:i]):
#                 AF = 0.02
#                 conditionPush(data["SAR_AF"], i, AF)
#                 conditionPush(data["SAR_Feature"], i, "descend")


""" double Checked
* BOLL: 计算布林线值, 数据更新到data
* reference: https://www.csai.cn/gupiao/1203359.html
* https://baike.baidu.com/item/%E5%B8%83%E6%9E%97%E7%BA%BF/3424486?fr=aladdin
* :param
*   N: 计算N日的MA, 默认值20
*   i: 序列号(当天所处天数)
* :return [中轨线, 上轨线, 下轨线]
"""


def BOLL(i, N=20, P=2):
    data = stock.data
    if i < N - 1:
        mid = 0
        top = 0
        bottom = 0
        conditionPush(data["BOLL_Mid"], i, mid)
        conditionPush(data["BOLL_Top"], i, top)
        conditionPush(data["BOLL_Bottom"], i, bottom)
        return [mid, top, bottom]
    else:
        mid = MA(i, N)
        arr = data["Close"][i - N + 1: i + 1]
        std = math.sqrt(variance(arr))  # 标准差
        top = mid + P * std
        bottom = mid - P * std
        conditionPush(data["BOLL_Mid"], i, mid)
        conditionPush(data["BOLL_Top"], i, top)
        conditionPush(data["BOLL_Bottom"], i, bottom)
        return [mid, top, bottom]


"""
* MIKE: 麦克指标
* :param
*   i: index
*   N: 12
* :return
* 
* code:
*     TYP:=(HIGH+LOW+CLOSE)/3;
*     LL:=LLV(LOW,N);
*     HH:=HHV(HIGH,N);
*     WR:TYP+(TYP-LL);
*     MR:TYP+(HH-LL);
*     SR:2*HH-LL;
*     WS:TYP-(HH-TYP);
*     MS:TYP-(HH-LL);
*     SS:2*LL-HH
"""


def MIKE(i, N=12):
    data = stock.data
    TYP, LL, HH, WR, MR, SR, WS, MS, SS = 0, 0, 0, 0, 0, 0, 0, 0, 0
    if i - N + 1 >= 0:
        TYP = (data["High"][i] + data["Low"][i] + data["Close"][i]) / 3
        LL = min(data["Low"][i - N + 1: i + 1])
        HH = max(data["High"][i - N + 1: i + 1])
        WR = TYP + (TYP - LL)
        MR = TYP + (HH - LL)
        SR = 2 * HH - LL
        WS = TYP - (HH - TYP)
        MS = TYP - (HH - LL)
        SS = 2 * LL - HH
    conditionPush(data["MIKE_WR"], i, WR)
    conditionPush(data["MIKE_MR"], i, MR)
    conditionPush(data["MIKE_SR"], i, SR)
    conditionPush(data["MIKE_WS"], i, WS)
    conditionPush(data["MIKE_MS"], i, MS)
    conditionPush(data["MIKE_SS"], i, SS)
    return [WR, MR, SR, WS, MS, SS]


"""
* 变动率指标ROC
* :param
*   i: 时间序号
*   N: N日前收盘价(12 recommended)
*   M: ROC M日移动平均(6 recommended)
* :return
* reference: https://baike.baidu.com/item/%E5%8F%98%E5%8A%A8%E7%8E%87%E6%8C%87%E6%A0%87/10387865?fr=aladdin
"""


def ROC(i, N=12, M=6):
    data = stock.data
    AX, BX, ROC, MAROC = 0, 0, 0, 0
    if i - N >= 0:
        AX = data["Close"][i] - data["Close"][i - N]
        BX = data["Close"][i - N]
        ROC = AX / BX * 100
    if i - 1 >= 0:
        MAROC = ROC * 2 / (M + 1) + data["ROC_MAROC"][i - 1] * (M - 1) / (M + 1)
    conditionPush(data["ROC_ROC"], i, ROC)
    conditionPush(data["ROC_MAROC"], i, MAROC)
    return [ROC, MAROC]


"""
* 动量指标MTM
* reference: https://baike.baidu.com/item/%E5%8A%A8%E9%87%8F%E6%8C%87%E6%A0%87/6453656
* :param
*   N: N日前收盘价(12 recommended)
*   M: MTM的M日移动平均(6 recommended)
* :return
*   [MTM, MTMMA]
"""


def MTM(i, N=12, M=6):
    data = stock.data
    MTM, SUM, MTMMA = 0, 0, 0
    if i - N >= 0:
        MTM = data["Close"][i] - data["Close"][i - N]
    conditionPush(data["MTM_MTM"], i, MTM)
    if i - 1 >= 0:
        MTMMA = data["MTM_MTMMA"][i - 1] * (M - 1) / (M + 1) + MTM * 2 / (M + 1)
    conditionPush(data["MTM_MTMMA"], i, MTMMA)
    return [MTM, MTMMA]


def generateIndex(workbook=None, sheet=None):
    data = stock.data
    # data["Open"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(2)[1:]]
    # data["Close"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(5)[1:]]
    # data["High"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(3)[1:]]
    # data["Low"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(4)[1:]]
    # data["Volume"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(8)[1:]]
    # data["Date"] = workbook.sheet_by_name(sheet).col_values(0)[1:]
    data["Open"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(0)[1:]]
    data["Close"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(1)[1:]]
    data["High"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(2)[1:]]
    data["Low"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(3)[1:]]
    data["Volume"] = [float(ele) for ele in workbook.sheet_by_name(sheet).col_values(4)[1:]]
    data["Date"] = [ele[0:4] + ele[5:7] + ele[8:] for ele in workbook.sheet_by_name(sheet).col_values(11)[1:]]
    # for i in range(1, row_num):
    #     # Date
    #     conditionGet(data["Date"], workbook.sheet_by_name(sheet).cell_value(i, 0))
    #     # Index_Code
    #     # data["Date"].append(float(workbook.sheet_by_name(sheet).cell_value(i, 1)))
    #     # Index_Chinese_Name_Full
    #     # data["Date"].append(float(workbook.sheet_by_name(sheet).cell_value(i, 2)))
    #     # Index_Chinese_Name
    #     # data["Date"].append(float(workbook.sheet_by_name(sheet).cell_value(i, 3)))
    #     # Index_English_Name_Full
    #     # data["Date"].append(float(workbook.sheet_by_name(sheet).cell_value(i, 4)))
    #     # Index_Chinese_Name
    #     # data["Date"].append(float(workbook.sheet_by_name(sheet).cell_value(i, 5)))
    #     # Open
    #     conditionGet(data["Open"], workbook.sheet_by_name(sheet).cell_value(i, 6))
    #     # High
    #     conditionGet(data["High"], workbook.sheet_by_name(sheet).cell_value(i, 7))
    #     # Low
    #     conditionGet(data["Low"], workbook.sheet_by_name(sheet).cell_value(i, 8))
    #     # Close
    #     conditionGet(data["Close"], workbook.sheet_by_name(sheet).cell_value(i, 9))
    #     # Change
    #     conditionGet(data["Change"], workbook.sheet_by_name(sheet).cell_value(i, 10))
    #     # Change_Per
    #     conditionGet(data["ChangePercent"], workbook.sheet_by_name(sheet).cell_value(i, 11))
    #     # Volume_M_Shares
    #     conditionGet(data["Volume"], workbook.sheet_by_name(sheet).cell_value(i, 12))
    #     # Turnover
    #     conditionGet(data["Turnover"], workbook.sheet_by_name(sheet).cell_value(i, 13))
    #     # ConsNumber
    #     # data["Date"].append(workbook.sheet_by_name(sheet_name).cell_value(i, 14))

    for i in range(0, len(data["Date"])):
        MACD(i)
        DMI(i)
        EXPMA(i)
        PSY(i)
        KDJ(i)
        WR(i)
        RSI(i)
        BIAS(i)
        VR(i)
        ARBR(i)
        ASI(i)
        # OBV(i)
        # SAR(i)
        BOLL(i)
        MIKE(i)
        ROC(i)
        MTM(i)
        # SAR(i)
    # print(data["SAR_SAR"])
    # for key in data.keys():
    # print(key, len(data[key]))
    # print(len(data.keys()))
    # df = pd.DataFrame(data)
    # print(df)
    # df.to_excel("DATA.xlsx")


if __name__ == "__main__":
    workbook = xlrd.open_workbook("../data/All_Stock_Info(hushen300).xls")
    sheet = workbook.sheet_names()[0]
    generateIndex(workbook, sheet)
    # pd.DataFrame(pd.DataFrame.from_dict(data, orient='index').values.T, columns=list(data.keys())).to_excel(
    #     "aaaaaaaaaaa.xlsx")
    # pd.DataFrame(data).to_excel("aaaaaa.xlsx")
