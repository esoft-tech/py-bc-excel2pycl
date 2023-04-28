import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell


class TestTokens(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'tests/{uuid.uuid4()}.py'
        cls.parser = Parser() \
            .set_excel_file_path('tests/tokens.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлик питоновский
        os.remove(cls.translation_file_path)

    def test_min_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 1, 2), Cell(0, 1, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MIN token are OK')

    def test_max_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 2, 2), Cell(0, 2, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MAX token are OK')

    def test_round_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 3, 2), Cell(0, 3, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ROUND token are OK')

    def test_if_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 4, 2), Cell(0, 4, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='IF token are OK')

    def test_vlookup_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 5, 2), Cell(0, 5, 3)])
        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='VLOOKUP token are OK')

    def test_OR_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 6, 2), Cell(0, 6, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='OR token are OK')

    def test_day_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 7, 2), Cell(0, 7, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DAY token are OK')

    def test_month_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 8, 2), Cell(0, 8, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MONTH token are OK')

    def test_year_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 9, 2), Cell(0, 9, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='YEAR token are OK')


if __name__ == '__main__':
    unittest.main()
