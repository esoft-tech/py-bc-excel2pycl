import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell


class TestCountBlankCcToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        cls.parser = Parser() \
            .set_excel_file_path('test/Excel2PyCl.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    # @classmethod
    # def tearDownClass(cls) -> None:
    #     # после выполнения всех тестов удаляем файлики
    #     os.remove(cls.translation_file_path)

    def test_standard(self):
        excepted_cell_value = 3
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 5, 0))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='test standart')

    def test_cell_concat(self):
        excepted_cell_value = 3
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 5, 1))

        cell_value2 = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 2, 1))

        print(cell_value2.value)

        cell_value2 = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 1, 1))

        print(cell_value2.value)

        self.assertEqual(cell_value.value, excepted_cell_value, msg='test sell concat')

    def test_cell_false(self):
        excepted_cell_value = 1
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 5, 2))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_rectangle(self):
        excepted_cell_value = 52
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 5, 9))

        self.assertEqual(cell_value.value, excepted_cell_value)

