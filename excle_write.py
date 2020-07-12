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
    def __init__(self, wb: xlsxwriter.Workbook):
        self.wb = wb
        self.ws = self.wb.add_worksheet("续保日报")
        self.style = Style(self.wb)
        self.nrow = 0
        self.ncol = 0

    def write_title(self):
        """
        写入表标题
        """

        date = (datetime.now() - timedelta(days=1)).strftime("%m月%d日")

        title = f"机构车险续保跟踪日报 ({date})"

        nrow = self.nrow
        ncol = self.ncol

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 8,
            data=title,
            cell_format=self.style.title,
        )

        self.ws.set_row(row=self.nrow, height=24)
        self.nrow += 1

        logging.info("表标题写入完成")

    def write_header(self):
        """
        写入表头
        """

        nrow = self.nrow
        ncol = self.ncol

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow + 1,
            last_col=ncol + 1,
            data="机构",
            cell_format=self.style.header,
        )

        ncol += 2

        month = (datetime.now() - timedelta(days=1)).strftime("%m")

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
            data="上日续保件数",
            cell_format=self.style.header,
        )

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol + 6,
            last_row=nrow + 1,
            last_col=ncol + 6,
            data=f"{int(month)}月已续保率",
            cell_format=self.style.header,
        )

        nrow += 1

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

        self.nrow += 2

    def write_data(self):
        """
        写入数据
        """

        nrow = self.nrow
        ncol = self.ncol

        three_list = {
            "曲靖": 0,
            "文山": 0,
            "大理": 0,
            "版纳": 0,
            "保山": 0,
            "怒江": 0,
            "昭通": 0,
            "昆明": 0,
        }

        three = Stats()

        three_month_renewable = three.month_renewable()
        three_month_renewed = three.month_renewed()

        month = (datetime.now() - timedelta(days=1)).strftime("%m")

        three_last_month_renewable = three.month_renewable(int(month) + 1)
        three_last_month_renewed = three.month_renewed(int(month) + 1)

        three_day_renewed = three.day_renewed()
        three_app_day_renewed = three.day_renewed(app=True)

        three_renewal_rate = three.renewal_rate()

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
                cell_format=self.style.number,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 3,
                number=three_month_renewed[key],
                cell_format=self.style.number,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 4,
                number=three_last_month_renewable[key],
                cell_format=self.style.number,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 5,
                number=three_last_month_renewed[key],
                cell_format=self.style.number,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 5,
                number=three_day_renewed[key],
                cell_format=self.style.number,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 5,
                number=three_app_day_renewed[key],
                cell_format=self.style.number,
            )
            self.ws.write_number(
                row=nrow,
                col=ncol + 5,
                number=three_renewal_rate[key],
                cell_format=self.style.number,
            )
            nrow += 1

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
