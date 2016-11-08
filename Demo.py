# coding=utf-8

import xlsxwriter

# data to be write in excel file
expenses = (
    ["Rent", 1000],
    ["Gas", 100],
    ["Food", 300],
    ["Gym", 50]
)


def write_data_2_xlsx():
    workbook = xlsxwriter.Workbook("Demo.xlsx")
    worksheet = workbook.add_worksheet()
    row, col = 0, 0
    for name, num in expenses:
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, num)
        row += 1
    # calculate the sum of all the row num
    worksheet.write(row, 1, "=SUM(B1:B4)")

    workbook.close()


write_data_2_xlsx()
