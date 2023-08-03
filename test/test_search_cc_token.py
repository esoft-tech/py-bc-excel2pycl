import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime


def create_test_table(file_name):
    wb = Workbook()
    ws = wb.active
    data = [
        ['р', 'имбирь', '=SEARCH(A1,B1)'],
        ['и', 'имбирь', '=SEARCH(A2,B2,2)'],
        ['и*ь',	'ИМБИРЬ', '=SEARCH(A3,B3)'],
        ['МБ?РЬ', 'МБ?РЬ', '=SEARCH(A4,B4)'],
        ['Река', 'Имбирь', '=SEARCH(A5,B5)'],
        ['м', 'имбирь', '=SEARCH(A6,B6,3)']
    ]

    for row in data:
        ws.append(row)

    tab = Table(displayName='base', ref='A1:E5')

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws.add_table(tab)

    wb.save(file_name)

class TestSearchCcToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        create_test_table('test/search.xlsx')
        cls.parser = Parser() \
            .set_excel_file_path('test/search.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        # os.remove(cls.translation_file_path)
        # os.remove('test/search.xlsx')
        pass

    def test_two_args(self):
        excepted_cell_value = 5
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 0))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token error (two args)')

    def test_three_args(self):
        excepted_cell_value = 4
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 1))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token error  (three args)')

    def test_with_asterisk(self):
        excepted_cell_value = 3
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 2))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token error (with asterisk)')

    def test_with_question(self):
        excepted_cell_value = 2
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 2, 3))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token error  (with question)')

    def test_not_found(self):
        excepted_cell_value = '#VALUE!'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 4))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token error  (not found)')


if __name__ == '__main__':
    unittest.main()
