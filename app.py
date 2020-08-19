import openpyxl as xl

wb = xl.load_workbook('transactions.xlsx')

sheet = wb['Sheet1']

for row in range(2, sheet.max_row+1):
    print(row)
    cell = sheet.cell(row,3)
    print(cell.value)

print('Done')
