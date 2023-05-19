from datetime import datetime
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def create_test_table(file_name):
    wb = Workbook()
    ws = wb.active

    data = [
        # ROUND - 2 строка
        ['=ROUND(A4, 2)', '=OR(B4, B5)', '=AND(C4, C5)', '=VLOOKUP(5, D4:D13, 1, FALSE())', '=IF(E6>E5, 44, 11)',
         '=SUM(F4:F13)', '=SUMIF(G4:G13, ">6")', '=AVERAGE(H4:H13)'],
        [10.24, '=TRUE()', '=FALSE()', 5, 44, 55, 34, 5.5],
        [10.239584, '', 0, 1, 1, 1, 1, 1],
        ['', 100, 1, 2, 2, 2, 2, 2],
        ['', '', '', 3, 3, 3, 3, 3],
        ['', '', '', 4, 4, 4, 4, 4],
        ['', '', '', 5, 5, 5, 5, 5],
        ['', '', '', 6, 6, 6, 6, 6],
        ['', '', '', 7, 7, 7, 7, 7],
        ['', '', '', 8, 8, 8, 8, 8],
        ['', '', '', 9, 9, 9, 9, 9],
        ['', '', '', 10, 10, 10, 10, 10],
    ]

    # add column headings. NB. these must be strings
    ws.append(["ROUND", "OR", "AND", "VLOOKUP", "IF", "SUM", "SUMIF", "AVERAGE"])
    for row in data:
        ws.append(row)

    tab = Table(displayName="base", ref="A1:E5")

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws.add_table(tab)

    ws_date = wb.create_sheet('date')

    data = [
        ['=DATE(A4, A5, A6)', '=DATE(B4, B5, B6)', '=DATE(C4, C5, C6)', '=DATE(D4, D5, D6)',
         '=DATE(E4, E5, E6)', '=DATE(F4, F5, F6)', '=DATE(G4, G5, G6)', '=DATE(H4, H5, H6)', '=DATE(I4, I5, I6)',
         '=DATE(J4, J5, J6)'],
        [datetime(2022, 10, 10), datetime(2000, 10, 10), '#NUM!', datetime(2023, 1, 1), datetime(2024, 6, 1),
         datetime(2021, 11, 1), datetime(2019, 6, 1), datetime(2022, 8, 28), datetime(2022, 3, 17),
         datetime(2022, 9, 29)],
        [2022, 100, 10000, 2022, 2022, 2022, 2022, 2022, 2022, 2022],
        [10, 10, 10, 13, 30, -1, -30, 5, 5, 10],
        [10, 10, 10, 1, 1, 1, 1, 120, -44, -1],
    ]

    # add column headings. NB. these must be strings
    ws_date.append(["DATE normal", "DATE year < 1900", "DATE year > 9999",
                    "DATE month > 12", "DATE month > 24", "DATE month < 1", "DATE month < -12",
                    "DATE days > 30", "DATE days < 0", "DATE days = -1"])
    for row in data:
        ws_date.append(row)

    wb.save(file_name)
