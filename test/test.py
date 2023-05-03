import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell
from worksheet_creator import create_test_table


class TestTokens(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        create_test_table('test/tokens.xlsx')
        cls.parser = Parser() \
            .set_excel_file_path('test/tokens.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)
        os.remove('test/tokens.xlsx')

    def test_round_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 0, 1), Cell(0, 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ROUND token are OK')

    def test_or_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 1, 1), Cell(0, 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='OR token are OK')

    def test_and_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 2, 1), Cell(0, 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='AND token are OK')

    def test_vlookup_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 3, 1), Cell(0, 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='VLOOKUP token are OK')

    def test_if_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 4, 1), Cell(0, 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='IF token are OK')

    def test_sum_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 5, 1), Cell(0, 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUM token are OK')

    def test_sumif_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 6, 1), Cell(0, 6, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIF token are OK')

    def test_average_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 7, 1), Cell(0, 7, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='AVERAGE token are OK')

    def test_min_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 8, 1), Cell(0, 8, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MIN token are OK')

    def test_max_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 9, 1), Cell(0, 9, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MAX token are OK')


if __name__ == '__main__':
    unittest.main()
