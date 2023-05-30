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

    def test_year_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 10, 1), Cell(0, 10, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='YEAR token are OK')

    def test_month_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 11, 1), Cell(0, 11, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MONTH token are OK')

    def test_day_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 12, 1), Cell(0, 12, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DAY token are OK')

    def test_iferror_when_error_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 13, 1), Cell(0, 13, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='IFERROR (when_error) token are OK')

    def test_iferror_condition_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 14, 1), Cell(0, 14, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='IFERROR (condition) token are OK')

    def test_date_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 0, 1), Cell('date', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (normal) token are OK')

    def test_date_lt_1900_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 1, 1), Cell('date', 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (year < 1900) token are OK')

    def test_date_bt_9999_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 2, 1), Cell('date', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (year > 9999) token are OK')

    def test_date_month_bt_12_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 3, 1), Cell('date', 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month > 12) token are OK')

    def test_date_month_bt_24_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 4, 1), Cell('date', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month > 24) token are OK')

    def test_date_month_lt_1_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 5, 1), Cell('date', 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month < 1) token are OK')

    def test_date_month_lt_24_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 6, 1), Cell('date', 6, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month < 24) token are OK')

    def test_date_day_120_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 7, 1), Cell('date', 7, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (days = 120) token are OK')

    def test_date_day_m_44_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 8, 1), Cell('date', 8, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (days = -44) token are OK')

    def test_date_day_eq_m1_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('date', 9, 1), Cell('date', 9, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (days = -1) token are OK')

    def test_datedif_d_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 0, 1), Cell('datedif', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF D token are OK')

    def test_datedif_m_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 1, 1), Cell('datedif', 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF M token are OK')

    def test_datedif_y_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 2, 1), Cell('datedif', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF Y token are OK')

    def test_datedif_md_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 3, 1), Cell('datedif', 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF MD token are OK')

    def test_datedif_ym_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 4, 1), Cell('datedif', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YM token are OK')

    def test_datedif_yd_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 5, 1), Cell('datedif', 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YD token are OK')

    def test_datedif_yd_month_lt_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 6, 1), Cell('datedif', 6, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YD token are OK')

    def test_datedif_yd_month_bt_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('datedif', 7, 1), Cell('datedif', 7, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YD token are OK')

    def test_eomonth_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 15, 1), Cell(0, 15, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='EOMONTH token are OK')

    def test_eomonth_negative_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell(0, 16, 1), Cell(0, 16, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='EOMONTH (N) token are OK')


if __name__ == '__main__':
    unittest.main()
