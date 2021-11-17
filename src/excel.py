from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


class Excel:
    def __init__(self, worksheets):
        self._data = worksheets['data']
        self._titles = worksheets['titles']

    @classmethod
    def parse(cls, path: str):
        worksheets_data = []
        worksheets_titles = []
        wb = load_workbook(filename=path)
        for worksheet in wb.worksheets:
            worksheets_titles.append(worksheet.title)
            worksheet_data = []
            last_column = len(list(worksheet.columns))
            last_row = len(list(worksheet.rows))
            for row_number in range(1, last_row + 1):
                rows_data = []
                for column_number in range(1, last_column + 1):
                    column_letter = get_column_letter(column_number)
                    rows_data.append(worksheet[f'{column_letter}{row_number}'].value)
                worksheet_data.append(rows_data)
            worksheets_data.append(worksheet_data)

        return cls({'data': worksheets_data, 'titles': worksheets_titles})
