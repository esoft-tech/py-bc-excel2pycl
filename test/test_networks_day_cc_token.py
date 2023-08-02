import datetime
import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def create_test_table(filename):
    wb = Workbook()
    ws = wb.active

    tab = Table(displayName='base', ref='A1:C5')

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws.add_table(tab)

    data = [
        [datetime.datetime(2023, 4, 1), datetime.datetime(2023, 5, 31), '=NETWORKDAYS(A1,B1)'],
        [datetime.datetime(2023, 5, 31), datetime.datetime(2023, 4, 1), '=NETWORKDAYS(A2,B2)'],
        [datetime.datetime(2023, 5, 1), datetime.datetime(2023, 8, 1), ''],
        [datetime.datetime(2023, 9, 1), 'ewewwewe', '=NETWORKDAYS(A1,B1,A3:B4)'],
        ['', '', '=NETWORKDAYS(A4,B4)']
    ]

    for row in data:
        ws.append(row)

    wb.save(filename)


class TestANetworkDaysCcToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        create_test_table('test/networks_days.xlsx')
        cls.parser = Parser() \
            .set_excel_file_path('test/networks_days.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        # os.remove(cls.translation_file_path)
        # os.remove('test/networks_days.xlsx')
        pass

    def test_two_args(self):
        excepted_cell_value = 43
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 0))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_three_args(self):
        excepted_cell_value = 42
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 3))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_incorrect_interval(self):
        excepted_cell_value = '#VALUE!'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 4))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_date_start_bigger_date_end(self):
        excepted_cell_value = -43
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 1))

        self.assertEqual(cell_value.value, excepted_cell_value)

