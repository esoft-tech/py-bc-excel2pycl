import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo



def create_test_table(file_name):
    wb = Workbook()
    ws = wb.active

    data = [
        ['', 0, 'dfdffd', '', '', '=COUNTBLANK(A1:E1)'],
        ['=""', '=""', '=A2&B2', 'wewe', 34, '=COUNTBLANK(A2:E2)'],
        ['', '', '', '', '', '=COUNTBLANK(A1:E2)']
    ]

    for row in data:
        ws.append(row)

    tab = Table(displayName='base', ref='A1:F5')

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws.add_table(tab)

    wb.save(file_name)


class TestCountBlankCcToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        create_test_table('test/count_blank.xlsx')
        cls.parser = Parser() \
            .set_excel_file_path('test/count_blank.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)
        os.remove('test/count_blank.xlsx')
        pass
    def test_standard(self):
        excepted_cell_value = 3
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 5, 0))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='test standart')

    def test_cell_concat(self):
        excepted_cell_value = 3
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 5, 1))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='test sell concat')

    def test_rectangle(self):
        excepted_cell_value = 6
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 5, 2))

        self.assertEqual(cell_value.value, excepted_cell_value)

