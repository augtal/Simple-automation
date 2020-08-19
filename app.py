import openpyxl as xl
from openpyxl.chart import BarChart, Reference


def process_workbook(filename):
    workbook = xl.load_workbook(filename)

    finalSheet = workbook.create_sheet(title="Final")
    topValues = ["Pirkejas", "Nr.", "Dokumento Data",
                 "Suma", "Nuolaidu suma", "Pirkejo skola"]

    finalSheet.delete_rows(1, 10)
    finalSheet.append(topValues)

    for sheet in workbook:
        if sheet.title == "Final":
            continue
        for row in sheet.iter_rows(min_row=1, max_col=sheet.max_column, max_row=10):
            values = []
            if row[0].value is None:
                continue
            elif row[0].value.find("Ex"):
                continue
            else:
                for cell in row:
                    if cell.value is None:
                        continue
                    else:
                        values.append(cell.value)
            values.insert(-1, None)
            finalSheet.append(values)

    workbook.save("Test3.xlsx")

    print('Done')


file_name = 'Test1.xlsx'
process_workbook(file_name)


"""
    for row in range(2, sheet.max_row+1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price

    values = Reference(sheet,
                       min_row=2, max_row=sheet.max_row,
                       min_col=4, max_col=4)

    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, 'e2')
"""

"""
        for row in range(1, sheet.max_row+1):
            cell = sheet.cell(row,1)
            if cell.value is None:
                continue

            if not cell.value.find("Ex"):
                sheet.unmerge_cells(start_row=row, start_column=1, end_row=row, end_column=sheet.max_column)
"""
