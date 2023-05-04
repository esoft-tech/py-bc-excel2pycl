from datetime import datetime
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def create_test_table(file_name):
    wb = Workbook()
    ws = wb.active

    data = [
        # ROUND - 2 строка
        ['=ROUND(A4, 2)', '=OR(B4, B5)', '=AND(C4, C5)', '=VLOOKUP(5, D4:D13, 1, FALSE())', '=IF(E6>E5, 44, 11)',
         '=SUM(F4:F13)', '=SUMIF(G4:G13, ">6")', '=AVERAGE(H4:H13)', '=DATE(I4, I5, I6)'],
        [10.24, '=TRUE()', '=FALSE()', 5, 44, 55, 34, 5.5, datetime(2022, 10, 10)],
        [10.239584, '', 0, 1, 1, 1, 1, 1, 2022],
        ['', 100, 1, 2, 2, 2, 2, 2, 10],
        ['', '', '', 3, 3, 3, 3, 3, 10],
        ['', '', '', 4, 4, 4, 4, 4],
        ['', '', '', 5, 5, 5, 5, 5],
        ['', '', '', 6, 6, 6, 6, 6],
        ['', '', '', 7, 7, 7, 7, 7],
        ['', '', '', 8, 8, 8, 8, 8],
        ['', '', '', 9, 9, 9, 9, 9],
        ['', '', '', 10, 10, 10, 10, 10],
    ]

    # add column headings. NB. these must be strings
    ws.append(["ROUND", "OR", "AND", "VLOOKUP", "IF", "SUM", "SUMIF", "AVERAGE", "DATE"])
    for row in data:
        ws.append(row)

    tab = Table(displayName="base", ref="A1:E5")

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws.add_table(tab)
    wb.save(file_name)
