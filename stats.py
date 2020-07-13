import logging
import sqlite3

from functools import lru_cache
from datetime import datetime
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
        self._weeknum = self._date.isocalendar()[1]

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

        self.four_list = {
            "春怡雅苑": 0,
            "香榭丽园": 0,
            "百大国际": 0,
            "春之城": 0,
            "宜良": 0,
            "东川": 0,
            "安宁": 0,
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
    def weeknum(self):
        """
        返回当前日期的周数
        """
        return self._weeknum

    @property
    def day(self):
        """
        返回今年的年份
        """
        return self._day

    @lru_cache(maxsize=32)
    def three_month_renewable(self, month: int = None):
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

        for key in self.three_list:
            if key not in data_dict:
                data_dict[key] = 0

        return data_dict

    def three_month_renewed(self, month: int = None):
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

        for key in self.three_list:
            if key not in data_dict:
                data_dict[key] = 0

        return data_dict

    def three_day_renewed(self, app: bool = False):
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
        last_date = f"{self.year}-{(month+2):02}-01"

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

        for key in self.three_list:
            if key not in data_dict:
                data_dict[key] = 0

        return data_dict

    def three_week_renewed(self, app: bool = False):
        """
        统计已续保清单的周数据
        """

        if app is True:
            sql_app = "AND ([已续保清单].[终端来源] = '0106移动展业(App)' OR \
                [已续保清单].[终端来源] = '0202APP')"
        else:
            sql_app = ""

        month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+2):02}-01"

        sql_str = f"SELECT \
            [中心支公司].[中心支公司简称], \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [中心支公司] ON [中心支公司].[中心支公司] = [已续保清单].[中心支公司] \
            JOIN [日期] ON [已续保清单].[投保确认日期] = [日期].[投保确认日期] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[中心支公司] = [可续保清单].[中支公司] \
            AND [日期].[周数] = '{self.weeknum}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            {sql_app} \
            GROUP  BY [已续保清单].[中心支公司];"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        for key in self.three_list:
            if key not in data_dict:
                data_dict[key] = 0

        return data_dict

    @lru_cache(maxsize=32)
    def four_month_renewable(self, month: int = None):
        """
        统计可续保数量
        """

        if month is None:
            month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+1):02}-01"

        sql_str = f"SELECT [机构].[机构简称], COUNT ([可续保清单].[保单号]) \
            FROM [可续保清单] \
            JOIN [机构] \
            ON [可续保清单].[出单机构] = [机构].[机构] \
            WHERE [可续保清单].[保险止期] >= '{first_date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            GROUP  BY [可续保清单].[出单机构]"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        four_dict = dict()

        for key in self.four_list:
            if key not in data_dict:
                four_dict[key] = 0
            else:
                four_dict[key] = data_dict[key]

        return four_dict

    def four_month_renewed(self, month: int = None):
        """
        统计已续保清单
        """

        if month is None:
            month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+1):02}-01"

        sql_str = f"SELECT \
            [机构].[机构简称], \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [机构] ON [机构].[机构] = [已续保清单].[机构] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[机构] = [可续保清单].[出单机构] \
            AND [已续保清单].[投保确认日期] < '{last_date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            GROUP  BY [已续保清单].[机构];"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        four_dict = dict()

        for key in self.four_list:
            if key not in data_dict:
                four_dict[key] = 0
            else:
                four_dict[key] = data_dict[key]

        return four_dict

    def four_day_renewed(self, app: bool = False):
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
        last_date = f"{self.year}-{(month+2):02}-01"

        sql_str = f"SELECT \
            [机构].[机构简称], \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [机构] ON [机构].[机构] = [已续保清单].[机构] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[机构] = [可续保清单].[出单机构] \
            AND [已续保清单].[投保确认日期] = '{self.date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            {sql_app} \
            GROUP  BY [已续保清单].[机构];"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        four_dict = dict()

        for key in self.four_list:
            if key not in data_dict:
                four_dict[key] = 0
            else:
                four_dict[key] = data_dict[key]

        return four_dict

    def four_week_renewed(self, app: bool = False):
        """
        统计已续保清单的周数据
        """

        if app is True:
            sql_app = "AND ([已续保清单].[终端来源] = '0106移动展业(App)' OR \
                [已续保清单].[终端来源] = '0202APP')"
        else:
            sql_app = ""

        month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+2):02}-01"

        sql_str = f"SELECT \
            [机构].[机构简称], \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [机构] ON [机构].[机构] = [已续保清单].[机构] \
            JOIN [日期] ON [已续保清单].[投保确认日期] = [日期].[投保确认日期] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[机构] = [可续保清单].[出单机构] \
            AND [日期].[周数] = '{self.weeknum}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            {sql_app} \
            GROUP  BY [已续保清单].[机构];"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        for value in data_list:
            data_dict[value[0]] = value[1]

        four_dict = dict()

        for key in self.four_list:
            if key not in data_dict:
                four_dict[key] = 0
            else:
                four_dict[key] = data_dict[key]

        return four_dict

    @lru_cache(maxsize=32)
    def two_month_renewable(self, month: int = None):
        """
        统计可续保数量
        """

        if month is None:
            month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+1):02}-01"

        sql_str = f"SELECT COUNT ([可续保清单].[保单号]) \
            FROM [可续保清单] \
            WHERE [可续保清单].[保险止期] >= '{first_date}' \
            AND [可续保清单].[保险止期] < '{last_date}'"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        data_dict["合计"] = data_list[0][0]

        return data_dict

    def two_month_renewed(self, month: int = None):
        """
        统计已续保清单
        """

        if month is None:
            month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+1):02}-01"

        sql_str = f"SELECT \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[投保确认日期] < '{last_date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}'"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        data_dict["合计"] = data_list[0][0]

        return data_dict

    def two_day_renewed(self, app: bool = False):
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
        last_date = f"{self.year}-{(month+2):02}-01"

        sql_str = f"SELECT \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [已续保清单].[投保确认日期] = '{self.date}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            {sql_app}"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        data_dict["合计"] = data_list[0][0]

        return data_dict

    def two_week_renewed(self, app: bool = False):
        """
        统计已续保清单的周数据
        """

        if app is True:
            sql_app = "AND ([已续保清单].[终端来源] = '0106移动展业(App)' OR \
                [已续保清单].[终端来源] = '0202APP')"
        else:
            sql_app = ""

        month = self.month

        first_date = f"{self.year}-{month:02}-01"
        last_date = f"{self.year}-{(month+2):02}-01"

        sql_str = f"SELECT \
            COUNT ([已续保清单].[续保单号]) AS [笔数] \
            FROM   [已续保清单] \
            JOIN [可续保清单] ON [已续保清单].[续保单号] = [可续保清单].[保单号] \
            JOIN [日期] ON [已续保清单].[投保确认日期] = [日期].[投保确认日期] \
            WHERE  [已续保清单].[续保单号] <> 'None' \
            AND [已续保清单].[保单笔数] = 1 \
            AND [日期].[周数] = '{self.weeknum}' \
            AND [可续保清单].[保险止期] < '{last_date}' \
            AND [可续保清单].[保险止期] >= '{first_date}' \
            {sql_app}"

        self.cur.execute(sql_str)
        data_list = self.cur.fetchall()

        data_dict = dict()

        data_dict["合计"] = data_list[0][0]

        return data_dict

    def renewal_rate(self, type: str):
        """
        统计已续保率
        """

        month = self.month

        if type == "three":
            renewadle = self.three_month_renewable(month=month)
            renewed = self.three_month_renewed(month=month)
        elif type == "four":
            renewadle = self.four_month_renewable(month=month)
            renewed = self.four_month_renewed(month=month)
        else:
            renewadle = self.two_month_renewable(month=month)
            renewed = self.two_month_renewed(month=month)

        data_dict = dict()

        for key in renewadle:
            data_dict[key] = renewed[key] / renewadle[key]

        return data_dict


if __name__ == "__main__":
    renew = Stats()
    data = dict()
    value = renew.two_month_renewable(month=7)
    print(value["合计"])
    # value = renew.four_month_renewable(month=7)
    # for key in renew.four_list:
    #     data[key] = value[key]
    # print(data)
    # value = renew.four_month_renewed(month=7)
    # for key in renew.four_list:
    #     data[key] = value[key]
    # print(data)
    # value = renew.day_renewed(app=True)
    # for key in renew.three_list:
    #     if key in value:
    #         data[key] = value[key]
    #     else:
    #         data[key] = 0
    # print(data)
    value = renew.renewal_rate(type="four")
    print(value)
    # value = renew.week_renewed(app=False)
    # for key in renew.three_list:
    #     if key in value:
    #         data[key] = value[key]
    #     else:
    #         data[key] = 0
    # print(data)

    # print(f"{renew.weeknum=}")
