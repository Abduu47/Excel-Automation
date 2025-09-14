# Excel-Automation
The following code gives you a simple structure to automate your excel files using openpyxl (Python library to manipulate excel files).
Code:

import openpyxl as xl
from openpyxl.chart import BarChart, Reference


def process_wb(filename):

    wb = xl.load_workbook(filename)
    sheet = wb['Sheet1']


    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price


    values = Reference(sheet, min_row=2, max_row=sheet.max_row, min_col=4, max_col=4)
    br = BarChart()
    br.add_data(values)
    sheet.add_chart(br, 'f2')

    wb.save(filename)

# Call process_wb funcition and give the file name as the argument.
# Can change the code according to your need.
