import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell
from .worksheet_creator import create_test_table


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

    @property
    def executor(self):
        return Executor().set_executed_class(class_file=self.translation_file_path)

    def test_round_token(self):
        cell_values = self.executor.get_cells([Cell(0, 0, 1), Cell(0, 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ROUND token are OK')

    def test_or_token(self):
        cell_values = self.executor.get_cells([Cell(0, 1, 1), Cell(0, 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='OR token are OK')

    def test_and_token(self):
        cell_values = self.executor.get_cells([Cell(0, 2, 1), Cell(0, 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='AND token are OK')

    def test_vlookup_token(self):
        cell_values = self.executor.get_cells([Cell(0, 3, 1), Cell(0, 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='VLOOKUP token are OK')

    def test_if_token(self):
        cell_values = self.executor.get_cells([Cell(0, 4, 1), Cell(0, 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='IF token are OK')

    def test_sum_token(self):
        cell_values = self.executor.get_cells([Cell(0, 5, 1), Cell(0, 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUM token are OK')

    def test_sumif_token(self):
        cell_values = self.executor.get_cells([Cell(0, 6, 1), Cell(0, 6, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIF token are OK')

    def test_average_token(self):
        cell_values = self.executor.get_cells([Cell(0, 7, 1), Cell(0, 7, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='AVERAGE token are OK')

    def test_min_token(self):
        cell_values = self.executor.get_cells([Cell(0, 8, 1), Cell(0, 8, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MIN token are OK')

    def test_max_token(self):
        cell_values = self.executor.get_cells([Cell(0, 9, 1), Cell(0, 9, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MAX token are OK')

    def test_year_token(self):
        cell_values = self.executor.get_cells([Cell(0, 10, 1), Cell(0, 10, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='YEAR token are OK')

    def test_month_token(self):
        cell_values = self.executor.get_cells([Cell(0, 11, 1), Cell(0, 11, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MONTH token are OK')

    def test_day_token(self):
        cell_values = self.executor.get_cells([Cell(0, 12, 1), Cell(0, 12, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DAY token are OK')

    def test_iferror_when_error_token(self):
        cell_values = self.executor.get_cells([Cell(0, 13, 1), Cell(0, 13, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='IFERROR (when_error) token are OK')

    def test_iferror_condition_token(self):
        cell_values = self.executor.get_cells([Cell(0, 14, 1), Cell(0, 14, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='IFERROR (condition) token are OK')

    def test_date_token(self):
        cell_values = self.executor.get_cells([Cell('date', 0, 1), Cell('date', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (normal) token are OK')

    def test_date_lt_1900_token(self):
        cell_values = self.executor.get_cells([Cell('date', 1, 1), Cell('date', 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (year < 1900) token are OK')

    def test_date_bt_9999_token(self):
        cell_values = self.executor.get_cells([Cell('date', 2, 1), Cell('date', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (year > 9999) token are OK')

    def test_date_month_bt_12_token(self):
        cell_values = self.executor.get_cells([Cell('date', 3, 1), Cell('date', 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month > 12) token are OK')

    def test_date_month_bt_24_token(self):
        cell_values = self.executor.get_cells([Cell('date', 4, 1), Cell('date', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month > 24) token are OK')

    def test_date_month_lt_1_token(self):
        cell_values = self.executor.get_cells([Cell('date', 5, 1), Cell('date', 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month < 1) token are OK')

    def test_date_month_lt_24_token(self):
        cell_values = self.executor.get_cells([Cell('date', 6, 1), Cell('date', 6, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (month < 24) token are OK')

    def test_date_day_120_token(self):
        cell_values = self.executor.get_cells([Cell('date', 7, 1), Cell('date', 7, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (days = 120) token are OK')

    def test_date_day_m_44_token(self):
        cell_values = self.executor.get_cells([Cell('date', 8, 1), Cell('date', 8, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (days = -44) token are OK')

    def test_date_day_eq_m1_token(self):
        cell_values = self.executor.get_cells([Cell('date', 9, 1), Cell('date', 9, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATE (days = -1) token are OK')

    def test_datedif_d_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 0, 1), Cell('datedif', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF D token are OK')

    def test_datedif_m_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 1, 1), Cell('datedif', 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF M token are OK')

    def test_datedif_y_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 2, 1), Cell('datedif', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF Y token are OK')

    def test_datedif_md_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 3, 1), Cell('datedif', 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF MD token are OK')

    def test_datedif_ym_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 4, 1), Cell('datedif', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YM token are OK')

    def test_datedif_yd_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 5, 1), Cell('datedif', 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YD token are OK')

    def test_datedif_yd_month_lt_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 6, 1), Cell('datedif', 6, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YD token are OK')

    def test_datedif_yd_month_bt_token(self):
        cell_values = self.executor.get_cells([Cell('datedif', 7, 1), Cell('datedif', 7, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='DATEDIF YD token are OK')

    def test_eomonth_token(self):
        cell_values = self.executor.get_cells([Cell(0, 15, 1), Cell(0, 15, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='EOMONTH token are OK')

    def test_eomonth_negative_token(self):
        cell_values = self.executor.get_cells([Cell(0, 16, 1), Cell(0, 16, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='EOMONTH (N) token are OK')

    def test_edate_months_plus_token(self):
        cell_values = self.executor.get_cells([Cell(0, 17, 1), Cell(0, 17, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='EDATE plus token are OK')

    def test_edate_months_minus_token(self):
        cell_values = self.executor.get_cells([Cell(0, 18, 1), Cell(0, 18, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='EDATE minus token are OK')

    def test_match_token(self):
        cell_values = self.executor.get_cells([Cell(0, 19, 1), Cell(0, 19, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MATCH token are OK')

    def test_xmatch_token(self):
        cell_values = self.executor.get_cells([Cell(0, 20, 1), Cell(0, 20, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='XMATCH token are OK')
    
    def test_left_normal_token(self):
        cell_values = self.executor.get_cells([Cell('left', 0, 1), Cell('left', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='LEFT token are OK')

    def test_left_zero_token(self):
        cell_values = self.executor.get_cells([Cell('left', 1, 1), Cell('left', 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='LEFT token are OK')
        
    def test_left_bt_len_token(self):
        cell_values = self.executor.get_cells([Cell('left', 2, 1), Cell('left', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='LEFT token are OK')
        
    def test_left_eq_len_token(self):
        cell_values = self.executor.get_cells([Cell('left', 3, 1), Cell('left', 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='LEFT token are OK')
        
    def test_left_1_arg_token(self):
        cell_values = self.executor.get_cells([Cell('left', 4, 1), Cell('left', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='LEFT token are OK')
        
    def test_left_neg_arg_token(self):
        cell_values = self.executor.get_cells([Cell('left', 5, 1), Cell('left', 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='LEFT token are OK')
        
    def test_right_normal_token(self):
        cell_values = self.executor.get_cells([Cell('right', 0, 1), Cell('right', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='right token are OK')

    def test_right_zero_token(self):
        cell_values = self.executor.get_cells([Cell('right', 1, 1), Cell('right', 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='right token are OK')
        
    def test_right_bt_len_token(self):
        cell_values = self.executor.get_cells([Cell('right', 2, 1), Cell('right', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='right token are OK')
        
    def test_right_eq_len_token(self):
        cell_values = self.executor.get_cells([Cell('right', 3, 1), Cell('right', 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='right token are OK')
        
    def test_right_1_arg_token(self):
        cell_values = self.executor.get_cells([Cell('right', 4, 1), Cell('right', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='right token are OK')

    def test_right_neg_arg_token(self):
        cell_values = self.executor.get_cells([Cell('right', 5, 1), Cell('right', 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='right token are OK')
        
    def test_mid_normal_token(self):
        cell_values = self.executor.get_cells([Cell('mid', 0, 1), Cell('mid', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MID token are OK')

    def test_mid_zero_token(self):
        cell_values = self.executor.get_cells([Cell('mid', 1, 1), Cell('mid', 1, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MID token are OK')
        
    def test_mid_bt_start_pos_token(self):
        cell_values = self.executor.get_cells([Cell('mid', 2, 1), Cell('mid', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MID token are OK')
        
    def test_mid_eq_start_pos_token(self):
        cell_values = self.executor.get_cells([Cell('mid', 3, 1), Cell('mid', 3, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MID token are OK')
        
    def test_mid_eq_len_token(self):
        cell_values = self.executor.get_cells([Cell('mid', 4, 1), Cell('mid', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MID token are OK')
        
    def test_mid_sec_arg_err_token(self):
        cell_values = self.executor.get_cells([Cell('mid', 5, 1), Cell('mid', 5, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MID token are OK')
        
    def test_mid_third_arg_err_token(self):
        cell_values = self.executor.get_cells([Cell('mid', 6, 1), Cell('mid', 6, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='MID token are OK')

    def test_countifs_text_condition(self):
        cell_values = self.executor.get_cells([Cell('countifs', 0, 2), Cell('countifs', 0, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNTIFS text condition down')

    def test_countifs_cell_condition(self):
        cell_values = self.executor.get_cells([Cell('countifs', 1, 2), Cell('countifs', 1, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNTIFS cell condition down')

    def test_countifs_lambda_condition(self):
        cell_values = self.executor.get_cells([Cell('countifs', 2, 2), Cell('countifs', 2, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNTIFS lambda condition down')

    def test_countifs_expression_condition(self):
        cell_values = self.executor.get_cells([Cell('countifs', 3, 2), Cell('countifs', 3, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNTIFS expression condition down')

    def test_countifs_any_text_condition(self):
        cell_values = self.executor.get_cells([Cell('countifs', 4, 2), Cell('countifs', 4, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNTIFS any text condition down')

    def test_countifs_pattern_text_condition(self):
        cell_values = self.executor.get_cells([Cell('countifs', 5, 2), Cell('countifs', 5, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNTIFS pattern text condition down')

    def test_index_row_with_row_index_only(self):
        cell_values = self.executor.get_cells([Cell('index', 0, 3), Cell('index', 1, 0)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='INDEX row with row index only down')

    def test_index_column_with_column_index_only(self):
        cell_values = self.executor.get_cells([Cell('index', 1, 3), Cell('index', 0, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='INDEX column with column index only down')

    def test_index_matrix_with_row_and_column_indexes(self):
        cell_values = self.executor.get_cells([Cell('index', 2, 3), Cell('index', 1, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value,
                         msg='INDEX matrix with row and column indexes down')

    def test_index_reference_form(self):
        cell_values = self.executor.get_cells([Cell('index', 3, 3), Cell('index', 2, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='INDEX reference form down')

    def test_index_array_return(self):

        cell_value = self.executor.get_cell(Cell('index', 0, 4))
        cell_values_array = [
            [cell.value for cell in self.executor.get_cells([Cell('index', 0, 0), Cell('index', 1, 0), Cell('index', 2, 0)])],
            [cell.value for cell in self.executor.get_cells([Cell('index', 0, 1), Cell('index', 1, 1), Cell('index', 2, 1)])],
            [cell.value for cell in self.executor.get_cells([Cell('index', 0, 2), Cell('index', 1, 2), Cell('index', 2, 2)])]
        ]

        self.assertEqual(cell_value.value, cell_values_array, msg='INDEX array return down')


if __name__ == '__main__':
    unittest.main()
