import logging
import sqlite3

from functools import lru_cache
from datetime import date, datetime
from datetime import timedelta


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
        self._date = datetime.now() - timedelta(days=1)
        self._year = self._date.year
        self._month = self._date.month
        self._day = self._date.day

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
    def date(self):
        """
        返回今年的年份
        """
        return self._date.strftime("%Y-%m-%d")

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
    def month_renewable(self, month: int = None):
        """
        统计可续保数量
        """

        if month is None:
            month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+1):02}-01"

        sql_str = f"SELECT [中心支公司].[中心支公司简称], COUNT ([可续保清单].[保单号]) \
            FROM [可续保清单] \
            JOIN [中心支公司] \
            ON [可续保清单].[中支公司] = [中心支公司].[中心支公司] \
            WHERE [可续保清单].[保险止期] >= '{first_date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            GROUP  BY [可续保清单].[中支公司]"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        return data_dict

    def month_renewed(self, month: int = None):
        """
        统计已续保清单
        """

        if month is None:
            month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+1):02}-01"

        sql_str = f"SELECT \
            [中心支公司].[中心支公司简称], \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [中心支公司] ON [中心支公司].[中心支公司] = [已续保清单].[中心支公司] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[中心支公司] = [可续保清单].[中支公司] \
            AND [已续保清单].[投保确认日期] < '{last_date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            GROUP  BY [已续保清单].[中心支公司];"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        return data_dict

    def day_renewed(self, app: bool = False):
        """
        统计已续保清单的日数据
        """

        if app is True:
            sql_app = "AND ([已续保清单].[终端来源] = '0106移动展业(App)' OR \
                [已续保清单].[终端来源] = '0202APP')"
        else:
            sql_app = ""

        month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+1):02}-01"

        sql_str = f"SELECT \
            [中心支公司].[中心支公司简称], \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [中心支公司] ON [中心支公司].[中心支公司] = [已续保清单].[中心支公司] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[中心支公司] = [可续保清单].[中支公司] \
            AND [已续保清单].[投保确认日期] = '{self.date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            {sql_app} \
            GROUP  BY [已续保清单].[中心支公司];"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        return data_dict

    def renewal_rate(self, month: int = None):
        """
        统计已续保率
        """

        month = self.month

        renewadle = self.month_renewable(month=month)
        renewed = self.month_renewed(month=month)



        data_dict = dict()

        for key in renewadle:
            data_dict[key] = renewed[key] / renewadle[key]

        return data_dict


if __name__ == "__main__":
    renew = Stats()
    data = dict()
    # value = renew.month_renewable(month=7)
    # for key in renew.three_list:
    #     data[key] = value[key]
    # print(data)
    # value = renew.month_renewed(month=7)
    # for key in renew.three_list:
    #     data[key] = value[key]
    # print(data)
    # value = renew.day_renewed(app=True)
    # for key in renew.three_list:
    #     if key in value:
    #         data[key] = value[key]
    #     else:
    #         data[key] = 0
    # print(data)
    value = renew.renewal_rate()
    for key in renew.three_list:
        if key in value:
            data[key] = value[key]
        else:
            data[key] = 0
    print(data)
