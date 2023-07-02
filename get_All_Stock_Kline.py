# 绘制所有股票的K线图, 选出有代表性的股票
import time
from pytdx.exhq import *
from pytdx.hq import *
import datetime
import pandas as pd


def code_handle(code):
    res = str(code)
    while len(res) < 6:
        res = '0' + res
    return res


def main():
    api = TdxHq_API().connect('119.147.212.81', 7709)
    # 处理股票代码
    content = pd.read_excel("../data/StockIndex.xlsx", sheet_name=0)
    stockCode = [ele[0] for ele in content[["code"]].values]
    stockName = [ele[0] for ele in content[["name"]].values]
    for i in range(0, len(stockCode)):
        stockCode[i] = code_handle(stockCode[i])
    # print(stockCode)
    with pd.ExcelWriter('All_Stock_Info.xlsx') as writer:
        for i in range(363, len(stockCode)):
            # 参数4 - 日K线
            print(stockCode[i], end=' ')
            print(str(i) + "/" + str(len(stockCode)))
            """
            :parameter
                category, market, stockCode, start, count
                类别, 市场代码, 证券代码, 开始时间(距离现在的天数), 请求K线数目
             0 5分钟K线  1 15分钟K线  2 30分钟K线  3 1小时K线  4 日K线 
             5 周K线  6 月K线  7 1分钟   8 1分钟K线  9 日K线  10 季K线  11 年K线
            """
            res = []
            res.extend(api.get_security_bars(9, 0, stockCode[i], 1600, 800))
            res.extend(api.get_security_bars(9, 0, stockCode[i], 800, 800))
            res.extend(api.get_security_bars(9, 0, stockCode[i], 0, 800))
            df = pd.DataFrame(res)
            df["name"] = stockName[i]
            df.to_excel(writer, sheet_name=str(stockCode[i]), index=False)
            time.sleep(2)


if __name__ == "__main__":
    main()