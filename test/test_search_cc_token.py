import os
import unittest
import uuid
from excel2pycl import Parser, Executor, Cell
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def create_test_table(file_name):
    wb = Workbook()
    ws = wb.active
    data = [
        ['р', 'имбирь', '=SEARCH(A1,B1)'],
        ['и', 'имбирь', '=SEARCH(A2,B2,2)'],
        ['и*ь',	'ИМБИРЬ', '=SEARCH(A3,B3)'],
        ['МБ?РЬ', 'имбирь', '=SEARCH(A4,B4)'],
        ['Река', 'Имбирь', '=SEARCH(A5,B5)'],
        ['м', 'имбирь', '=SEARCH(A6,B6,3)'],
        [r'\d+', '12356', '=SEARCH(A7,B7)'],
        ['?ткрыт*р', 'эти открытые двери', '=SEARCH(A8,B8)'],
        ['п?чему же~?', 'Почему, почему же?', '=SEARCH(A9,B9)'],
        ['П*очему', 'Почему, почему же?', '=SEARCH(A10,B10)'],
        ['П?очему', 'почему', '=SEARCH(A11,B11)'],
        ['П?че\dу', 'поче5у', '=SEARCH(A12,B12)'],
        ['?мбирь', 'имбирь', '=SEARCH(A13,B13)'],
        ['* ?мбирь', 'консервированный имбирь', '=SEARCH(A14,B14)']

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
        os.remove(cls.translation_file_path)
        # os.remove('test/search.xlsx')

    def test_two_args(self):
        excepted_cell_value = 5
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 0))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_three_args(self):
        excepted_cell_value = 4
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 1))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_with_asterisk(self):
        excepted_cell_value = 1
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 2))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_with_question(self):
        excepted_cell_value = 2
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 3))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_not_found(self):
        excepted_cell_value = '#VALUE!'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 4))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_not_found_three_args(self):
        excepted_cell_value = '#VALUE!'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 5))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_reg(self):
            excepted_cell_value = '#VALUE!'
            cell_value = Executor() \
                .set_executed_class(class_file=self.translation_file_path) \
                .get_cell(Cell(0, 2, 6))

            self.assertEqual(cell_value.value, excepted_cell_value)

    def test_asterisk_and_question(self):
        excepted_cell_value = 5
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 7))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_with_tilda(self):
        excepted_cell_value = 9
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 8))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_asterisk_empty(self):
        excepted_cell_value = 1
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 9))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_question_empty(self):
        excepted_cell_value = '#VALUE!'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 10))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_reg2(self):
        excepted_cell_value = '#VALUE!'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 11))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_start_with_question(self):
        excepted_cell_value = 1
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 12))
        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_start_with_asterisk(self):
        excepted_cell_value = 1
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 2, 13))
        self.assertEqual(cell_value.value, excepted_cell_value)



if __name__ == '__main__':
    unittest.main()
