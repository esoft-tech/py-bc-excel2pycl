import datetime
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
        [93, '=IFS(A1>89,"A",A1>79,"B",A1>69,"C",A1>59,"D")'],
        [89, '=IFS(A2>89,"A",A2>79,"B",A2>69,"C",A2>59,"D")'],
        [71, '=IFS(A3>89,"A",A3>79,"B",A3>69,"C",A3>59,"D")'],
        [60, '=IFS(A4>89,"A",A4>79,"B",A4>69,"C",A4>59,"D")'],
        [58, '=IFS(A5>89;"A",A5>79,"B",A5>69,"C",A5>59,"D")'],
        [50, '=IFS(50>"нет","больше",50<"нет","меньше")'],
        ["120", "3"],
        ['', '=IFS(A7>B7,"верно",A7<B7,"неверно")'],
        ['', '=IFS(B1>B2, "B1>B2", B1<=B2, "B1<=B2")']
    ]

    for row in data:
        ws.append(row)

    tab = Table(displayName='base', ref='A1:B10')

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws.add_table(tab)

    wb.save(file_name)


class TestIfsToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        create_test_table('test/ifs.xlsx')
        cls.parser = Parser() \
            .set_excel_file_path('test/ifs.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)
        os.remove('test/ifs.xlsx')

    def test_true_in_first(self):
        excepted_cell_value = 'A'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 0))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='test_true_in_first')

    def test_true_in_last(self):
        excepted_cell_value = 'D'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 3))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='test sell concat')

    def test_true_in_middle(self):
        excepted_cell_value = 'C'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 2))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_not_true(self):
        excepted_cell_value = '#N/A'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 4))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_compare_word_and_number(self):
        excepted_cell_value = 'меньше'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 5))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_compare_str_as_number(self):
        excepted_cell_value = 'верно'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 7))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_compare_sell_and_sell(self):
        excepted_cell_value = 'B1<=B2'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 8))

        self.assertEqual(cell_value.value, excepted_cell_value)