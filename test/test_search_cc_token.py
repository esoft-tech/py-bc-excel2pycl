import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell


class TestSearchCcToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        cls.parser = Parser() \
            .set_excel_file_path('test/Excel2PyCl.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)

    def test_two_args(self):
        excepted_cell_value = 5
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 3, 0))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token is okay (two args)')

    # def test_three_args(self):
    #     excepted_cell_value = 4
    #     cell_value = Executor() \
    #         .set_executed_class(class_file=self.translation_file_path) \
    #         .get_cell(Cell(1, 3, 1))
    #
    #     self.assertEqual(cell_value.value, excepted_cell_value, msg='search token is okay (three args)')

    def test_with_asterisk(self):
        excepted_cell_value = 3
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 3, 2))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token is okay (with asterisk)')

    def test_with_question(self):
        excepted_cell_value = 2
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 3, 3))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token is okay (with question)')

    def test_not_found(self):
        excepted_cell_value = '#VALUE!'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(1, 3, 4))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='search token is okay (with question)')


if __name__ == '__main__':
    unittest.main()
