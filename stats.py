import logging
import sqlite3

from functools import lru_cache
from datetime import datetime


logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO, format=" %(asctime)s | %(levelname)s | %(message)s"
)


class Stats(object):
    """
    统计续保数据
    """

    def __init__(self):
        self._year = datetime.now().year
        self._month = datetime.now().month
        self._day = datetime.now().day

        self.conn = sqlite3.connect(r"Data\data.db")
        self.cur = self.conn.cursor()

        sql_str = "ATTACH DATABASE 'Data\\可续保清单.db' AS [可续保清单]"
        self.cur.execute(sql_str)
        sql_str = "ATTACH DATABASE 'Data\\已续保清单.db' AS [已续保清单]"
        self.cur.execute(sql_str)

        self.three_list = {
            "曲靖": 0,
            "文山": 0,
            "大理": 0,
            "版纳": 0,
            "保山": 0,
            "怒江": 0,
            "昭通": 0,
            "昆明": 0,
        }

    @property
    def year(self):
        """
        返回今年的年份
        """
        return self._year

    @property
    def month(self):
        """
        返回今年的年份
        """
        return self._month

    @property
    def day(self):
        """
        返回今年的年份
        """
        return self._day

    @lru_cache(maxsize=32)
    def Renewable(self, month: int = None):
        """
        统计可续保数量
        """

        if month is None:
            month = self.month

        sql_str = f"SELECT [可续保清单].[中支公司], COUNT([可续保清单].[续保单号]) AS [笔数] \
            FROM [已续保清单] \
            JOIN [可续保清单] \
            ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [日期] \
            ON [可续保清单].[保险止期] = [日期].[投保确认日期] \
            WHERE [可续保清单].[续保单号] <> 'None' \
            AND [可续保清单].[机构] = [可续保清单].[出单机构] \
            AND [日期].[月份] = '{month}' \
            GROUP BY [可续保清单].[中支公司]"

        sql_str = f"SELECT [中心支公司].[中心支公司简称], COUNT ([可续保清单].[保单号]) \
            FROM [可续保清单] \
            JOIN [日期] \
            ON [可续保清单].[保险止期] = [日期].[投保确认日期] \
            JOIN [中心支公司] \
            ON [可续保清单].[中支公司] = [中心支公司].[中心支公司] \
            WHERE [日期].月份 = '{month}' \
            GROUP  BY [可续保清单].[中支公司]"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        return data_dict


if __name__ == "__main__":
    renew = Stats()
    value = renew.Renewable(month=6)
    data = dict()
    for key in renew.three_list:
        data[key] = value[key]
    print(data)
