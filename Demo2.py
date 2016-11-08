# coding=utf-8

import xlsxwriter

expenses = (
    ["Rent", 1000],
    ["Gas", 100],
    ["Food", 300],
    ["Gym", 50]
)


def generate_style_xlsx_file():
    workbook = xlsxwriter.Workbook("Demo2.xlsx")
    worksheet = workbook.add_worksheet("FistSheet")

    # create two stylesheet
    bold_style = workbook.add_format({"bold": True})
    dollar_style = workbook.add_format({"num_format": "$#,##0"})

    row, col = 1, 0
    worksheet.write("A1", "Item", bold_style)
    worksheet.write("B1", "Cost", bold_style)

    for name, cost in expenses:
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, cost, dollar_style)
        row += 1
    worksheet.write(row, 0, "Total", bold_style)
    worksheet.write(row, 1, "=SUM(B2:B5)", bold_style)


generate_style_xlsx_file()
