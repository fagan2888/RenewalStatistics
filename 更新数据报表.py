import xlsxwriter

from excle_write import Excel_Write
from update import update


update("已续保清单")

wb = xlsxwriter.Workbook("续保统计表.xlsx")
excel = Excel_Write(wb)

excel.make_form()

wb.close()
