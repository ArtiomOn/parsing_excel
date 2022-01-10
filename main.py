import xlrd
from xlutils.copy import copy

rb = xlrd.open_workbook('name.xls', formatting_info=True)
r_sheet = rb.sheet_by_index(0)
r = r_sheet.nrows
wb = copy(rb)
current_sheet = wb.get_sheet(0)

for sheet in rb.sheets():
    for rowidx in range(r):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            if cell.value == "cell_value":
                try:
                    coming_number_in_stock_red = float(
                        str(sheet.cell(rowidx, colidx + 3)).partition('number:')[2].partition('(X')[0])
                    coming_money_red = float(
                        str(sheet.cell(rowidx, colidx + 4)).partition('number:')[2].partition('(X')[0])
                    consumption_number_in_stock_red = float(
                        str(sheet.cell(rowidx, colidx + 5)).partition('number:')[2].partition('(X')[0])
                    consumption_money_red = float(
                        str(sheet.cell(rowidx, colidx + 6)).partition('number:')[2].partition('(X')[0])

                    coming_number_in_stock_white = float(
                        str(sheet.cell(rowidx - 1, colidx + 3)).partition('number:')[2].partition('(X')[0])
                    coming_money_white = float(
                        str(sheet.cell(rowidx - 1, colidx + 4)).partition('number:')[2].partition('(X')[0])
                    consumption_number_in_stock_white = float(
                        str(sheet.cell(rowidx - 1, colidx + 5)).partition('number:')[2].partition('(X')[0])
                    consumption_money_white = float(
                        str(sheet.cell(rowidx - 1, colidx + 6)).partition('number:')[2].partition('(X')[0])

                    coming_number_in_stock_done = coming_number_in_stock_white - coming_number_in_stock_red
                    coming_money_done = coming_money_white - coming_money_red
                    consumption_number_in_stock_done = \
                        consumption_number_in_stock_white - consumption_number_in_stock_red
                    consumption_money_done = consumption_money_white - consumption_money_red

                    current_sheet.write(rowidx - 1, colidx + 3, coming_number_in_stock_done)
                    current_sheet.write(rowidx - 1, colidx + 4, coming_money_done)
                    current_sheet.write(rowidx - 1, colidx + 5, consumption_number_in_stock_done)
                    current_sheet.write(rowidx - 1, colidx + 6, consumption_money_done)

                    wb.save('name.xls')

                except ValueError:
                    continue
