# coding=utf-8
from datetime import datetime
import xlsxwriter

expenses = (
    ['Rent', '2013-01-13', 1000],
    ['Gas', '2013-01-14', 100],
    ['Food', '2013-01-16', 300],
    ['Gym', '2013-01-20', 50],
)


def generate_more_style_xlsx():
    workbook = xlsxwriter.Workbook("Demo3.xlsx")
    worksheet = workbook.add_worksheet("FistSheet")
    # generate some format
    bold_format = workbook.add_format({"bold": 1})
    dollar_format = workbook.add_format({"num_format": "$#,##0"})
    date_format = workbook.add_format({"num_format": "mmmm d yyyy"})

    # adjust the column width
    worksheet.set_column("B:B", 30)

    # generate the fist row
    worksheet.write(0, 0, "Item", bold_format)
    worksheet.write(0, 1, "Date", bold_format)
    worksheet.write(0, 2, "Cost", bold_format)

    # generate all data row
    row, col = 1, 0
    for name, date_str, cost in expenses:
        date = datetime.strptime(date_str, "%Y-%m-%d")
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, date, date_format)
        worksheet.write(row, col + 2, cost, dollar_format)
        row += 1

    # generate the last row
    worksheet.write(row, 0, "Total", bold_format)
    worksheet.write(row, 2, "=SUM(B2:B5)", dollar_format)

    # close file
    workbook.close()


generate_more_style_xlsx()
