import xlsxwriter
import logging

from datetime import datetime
from datetime import timedelta

from style import Style
from stats import Stats

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO, format=" %(asctime)s | %(levelname)s | %(message)s"
)


class Excel_Write(object):
    """
    Excel表格写入类，用于编写Excel表格
    """
    def __init__(self, wb: xlsxwriter.Workbook):
        self.wb = wb
        self.ws = self.wb.add_worksheet("续保日报表")
        self.style = Style(self.wb)
        self.nrow = 0
        self.ncol = 0

    def write_title(self):
        """
        写入表标题
        """

        # 获取前一日日期，日报表的数据截止至上一日
        date = (datetime.now() - timedelta(days=1)).strftime("%m月%d日")

        # 表标题
        title = f"机构车险续保跟踪日报 ({date})"

        # 行计数器
        nrow = self.nrow
        # 列计数器
        ncol = self.ncol

        # 写入表标题 表宽10列
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 10,
            data=title,
            cell_format=self.style.title,
        )

        # 设置行高，行高为两倍字体大小
        self.ws.set_row(row=self.nrow, height=24)
        self.nrow += 1

        logging.info("表标题写入完成")

    def write_header(self):
        """
        写入表头
        """

        # 行计数器
        nrow = self.nrow
        # 列计数器
        ncol = self.ncol

        # 写入表头的第一列，第一列占两行两列
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow + 1,
            last_col=ncol + 1,
            data="机构",
            cell_format=self.style.header,
        )

        ncol += 2

        # 获取当前月份
        month = (datetime.now() - timedelta(days=1)).strftime("%m")

        # 写入表头第一行的信息
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 1,
            data=f"{int(month)}月",
            cell_format=self.style.header,
        )

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol + 2,
            last_row=nrow,
            last_col=ncol + 3,
            data=f"{int(month) + 1}月",
            cell_format=self.style.header,
        )

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol + 4,
            last_row=nrow,
            last_col=ncol + 5,
            data="本周续保件数",
            cell_format=self.style.header,
        )

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol + 6,
            last_row=nrow,
            last_col=ncol + 7,
            data="上日续保件数",
            cell_format=self.style.header,
        )

        # 写入表头的最后一列，最后一列占两行两列
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol + 8,
            last_row=nrow + 1,
            last_col=ncol + 8,
            data=f"{int(month)}月已续保率",
            cell_format=self.style.header,
        )

        nrow += 1

        # 写入表头第二行信息
        self.ws.write_string(
            row=nrow, col=ncol, string="可续保", cell_format=self.style.header
        )
        ncol += 1
        self.ws.write_string(
            row=nrow, col=ncol, string="已续保", cell_format=self.style.header
        )
        ncol += 1
        self.ws.write_string(
            row=nrow, col=ncol, string="可续保", cell_format=self.style.header
        )
        ncol += 1
        self.ws.write_string(
            row=nrow, col=ncol, string="已续保", cell_format=self.style.header
        )
        ncol += 1
        self.ws.write_string(
            row=nrow, col=ncol, string="续保小计", cell_format=self.style.header
        )
        ncol += 1
        self.ws.write_string(
            row=nrow, col=ncol, string="APP续保", cell_format=self.style.header
        )
        ncol += 1
        self.ws.write_string(
            row=nrow, col=ncol, string="续保小计", cell_format=self.style.header
        )
        ncol += 1
        self.ws.write_string(
            row=nrow, col=ncol, string="APP续保", cell_format=self.style.header
        )

        self.nrow += 2

        logging.info("表头写入完成")

    def write_data(self):
        """
        写入数据
        """

        nrow = self.nrow
        ncol = self.ncol

        # 三级机构名单
        three_list = {
            "曲靖": 0,
            "文山": 0,
            "大理": 0,
            "版纳": 0,
            "保山": 0,
            "怒江": 0,
            "昭通": 0,
        }

        # 四级机构名单
        four_list = {
            "春怡雅苑": 0,
            "香榭丽园": 0,
            "百大国际": 0,
            "春之城": 0,
            "宜良": 0,
            "东川": 0,
            "安宁": 0,
        }

        # 获取当前月份
        month = int((datetime.now() - timedelta(days=1)).strftime("%m"))

        # 获取三级机构相应数据
        three = Stats()

        three_month_renewable = three.three_month_renewable()
        three_month_renewed = three.three_month_renewed()
        three_last_month_renewable = three.three_month_renewable(month + 1)
        three_last_month_renewed = three.three_month_renewed(month + 1)
        three_week_renewed = three.three_week_renewed()
        three_app_week_renewed = three.three_week_renewed(app=True)
        three_day_renewed = three.three_day_renewed()
        three_app_day_renewed = three.three_day_renewed(app=True)
        three_renewal_rate = three.renewal_rate(type="three")

        # 写入三级机构信息
        for key in three_list:
            self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow,
                last_col=ncol + 1,
                data=key,
                cell_format=self.style.header,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 2,
                number=three_month_renewable[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 3,
                number=three_month_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 4,
                number=three_last_month_renewable[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 5,
                number=three_last_month_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 6,
                number=three_week_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 7,
                number=three_app_week_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 8,
                number=three_day_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 9,
                number=three_app_day_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 10,
                number=three_renewal_rate[key],
                cell_format=self.style.percent,
            )
            nrow += 1

        # 获取四级机构相关数据
        four = Stats()

        four_month_renewable = four.four_month_renewable()
        four_month_renewed = four.four_month_renewed()
        four_last_month_renewable = four.four_month_renewable(month + 1)
        four_last_month_renewed = four.four_month_renewed(month + 1)
        four_week_renewed = four.four_week_renewed()
        four_app_week_renewed = four.four_week_renewed(app=True)
        four_day_renewed = four.four_day_renewed()
        four_app_day_renewed = four.four_day_renewed(app=True)
        four_renewal_rate = four.renewal_rate(type="four")

        # 写入四级机构数据
        self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow + 7,
                last_col=ncol,
                data="昆明地区",
                cell_format=self.style.header,
            )

        for key in four_list:
            self.ws.write_string(
                row=nrow,
                col=ncol + 1,
                string=key,
                cell_format=self.style.header,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 2,
                number=four_month_renewable[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 3,
                number=four_month_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 4,
                number=four_last_month_renewable[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 5,
                number=four_last_month_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 6,
                number=four_week_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 7,
                number=four_app_week_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 8,
                number=four_day_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 9,
                number=four_app_day_renewed[key],
                cell_format=self.style.number_0,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 10,
                number=four_renewal_rate[key],
                cell_format=self.style.percent,
            )
            nrow += 1

        # 单独写入昆明地区的小计
        self.ws.write_string(
                row=nrow,
                col=ncol + 1,
                string="小计",
                cell_format=self.style.header,
            )
        self.ws.write_number(
            row=nrow,
            col=ncol + 2,
            number=three_month_renewable["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 3,
            number=three_month_renewed["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 4,
            number=three_last_month_renewable["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 5,
            number=three_last_month_renewed["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 6,
            number=three_week_renewed["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 7,
            number=three_app_week_renewed["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 8,
            number=three_day_renewed["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 9,
            number=three_app_day_renewed["昆明"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 10,
            number=three_renewal_rate["昆明"],
            cell_format=self.style.percent,
        )
        nrow += 1

        # 获取分公司的整体信息
        two = Stats()

        two_month_renewable = two.two_month_renewable()
        two_month_renewed = two.two_month_renewed()
        two_last_month_renewable = two.two_month_renewable(month + 1)
        two_last_month_renewed = two.two_month_renewed(month + 1)
        two_week_renewed = two.two_week_renewed()
        two_app_week_renewed = two.two_week_renewed(app=True)
        two_day_renewed = two.two_day_renewed()
        two_app_day_renewed = two.two_day_renewed(app=True)
        two_renewal_rate = two.renewal_rate(type="two")

        # 将分公司整体信息写入表的最后一行
        self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow,
                last_col=ncol + 1,
                data="合计",
                cell_format=self.style.header,
            )
        self.ws.write_number(
            row=nrow,
            col=ncol + 2,
            number=two_month_renewable["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 3,
            number=two_month_renewed["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 4,
            number=two_last_month_renewable["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 5,
            number=two_last_month_renewed["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 6,
            number=two_week_renewed["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 7,
            number=two_app_week_renewed["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 8,
            number=two_day_renewed["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 9,
            number=two_app_day_renewed["合计"],
            cell_format=self.style.number_0,
        )
        self.ws.write_number(
            row=nrow,
            col=ncol + 10,
            number=two_renewal_rate["合计"],
            cell_format=self.style.percent,
        )

        logging.info("表数据写入完成")

    def make_form(self):
        """
        制作表格
        """

        self.write_title()
        self.write_header()
        self.write_data()


if __name__ == "__main__":
    wb = xlsxwriter.Workbook("test.xlsx")
    excel = Excel_Write(wb)

    excel.make_form()

    wb.close()
