# how to process spreadsheets
import openpyxl as xl
from openpyxl.chart import BarChart, Reference

wb = xl.load_workbook('transactions.xlsx')
sheet = wb['Sheet1']

# get the coordinates of the cell
# column then row
cell = sheet['a1']

# or use the .cell() method of the sheet object
# pass in the row and column
cell = sheet.cell(1,1)

# shows transaction_id in console
print(cell.value)

# to determine how many rows in the spreadsheet
# shows 4, so we need to generate a for loop that would generate the numbers 1-4
print(sheet.max_row)

for row in range(2, sheet.max_row + 1):
    price_cell = sheet.cell(row, 3)
    corrected_price = price_cell.value * 0.9
    corrected_price_cell = sheet.cell(row, 4)
    corrected_price_cell.value = corrected_price

wb.save('transactions2.xlsx')

