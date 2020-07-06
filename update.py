import logging
import sqlite3

from datetime import datetime
from openpyxl import load_workbook

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO, format=" %(asctime)s | %(levelname)s | %(message)s"
)


def update(table_name):
    logging.info("开始更新数据库")
    data_path = r'Data\已续保清单.db'
    if table_name == "可续保清单":
        data_path = r'Data\可续保清单.db'
    conn = sqlite3.connect(data_path)
    logging.info('数据库连接成功')
    cur = conn.cursor()

    # 清空原数据库数据
    str_sql = f"DELETE FROM [{table_name}]"
    cur.execute(str_sql)
    conn.commit()
    logging.info("数据库数据清空完毕")

    # 读入Excel表格数据
    wb = load_workbook(f'{table_name}.xlsx')
    ws = wb['page']

    max_col = ws.max_column
    max_row = ws.max_row

    logging.info("Excel 文件读入成功")
    logging.info(f"需要导入{ws.max_row}条数据")

    title_row = ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=max_col)

    titles = []

    for title_tuple in title_row:
        for key in title_tuple:
            titles.append(key.value)

    begin_row = 2

    for row in ws.iter_rows(min_row=begin_row, max_row=max_row, min_col=1,
                            max_col=max_col):
        str_sql = f"INSERT INTO '{table_name}' VALUES ("
        for key in row:
            str_sql += f"'{key.value}', "
        str_sql = str_sql[:-2] + ')'
        cur.execute(str_sql)

    logging.info('数据写入数据库完成')

    conn.commit()

    logging.info("数据库事务提交完成")
    logging.info("数据库更新操作完成")

    cur.close()
    conn.close()

    print("-" * 60)


if __name__ == "__main__":
    update("可续保清单")
    update("已续保清单")
