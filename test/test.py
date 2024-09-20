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

    def test_count_token(self):
        cell_values = self.executor.get_cells([Cell('count', 0, 2), Cell('count', 0, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNT token with range argument down')

    def test_count_token_single_cell_arg(self):
        cell_values = self.executor.get_cells([Cell('count', 0, 3), Cell('count', 1, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNT token with single cell argument down')

    def test_count_token_num_and_string_digits(self):
        cell_values = self.executor.get_cells([Cell('count', 0, 4), Cell('count', 2, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNT token with string digits down')

    def test_count_token_range_and_arg_sequence(self):
        cell_values = self.executor.get_cells([Cell('count', 0, 5), Cell('count', 3, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNT token with argument sequence down')

    def test_count_token_range_and_arg_sequence_with_bool_and_string_digit(self):
        cell_values = self.executor.get_cells([Cell('count', 0, 6), Cell('count', 4, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNT token with bool in argument down')

    def test_count_with_date(self):
        cell_values = self.executor.get_cells([Cell('count', 0, 7), Cell('count', 5, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='COUNT token with date in array down')

    def test_address_absolute_token(self):
        cell_values = self.executor.get_cells([Cell('address', 0, 2), Cell('address', 0, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS absolute down')

    def test_address_absolute_huge_token(self):
        cell_values = self.executor.get_cells([Cell('address', 8, 2), Cell('address', 8, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS absolute huge down')

    def test_address_relative_col_token(self):
        cell_values = self.executor.get_cells([Cell('address', 1, 2), Cell('address', 1, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS absolute row, relative col down')

    def test_address_relative_row_token(self):
        cell_values = self.executor.get_cells([Cell('address', 2, 2), Cell('address', 2, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS absolute col, relative row down')

    def test_address_relative_token(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('address', 3, 2), Cell('address', 3, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS relative down')

    def test_address_absolute_strict_token(self):
        cell_values = self.executor.get_cells([Cell('address', 4, 2), Cell('address', 4, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS absolute strict down')

    def test_address_rc_type_col_relative_token(self):
        cell_values = self.executor.get_cells([Cell('address', 5, 2), Cell('address', 5, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS RC type col relative down')

    def test_address_rc_type_col_relative_link_sheet_token(self):
        cell_values = self.executor.get_cells([Cell('address', 6, 2), Cell('address', 6, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS RC type col relative link sheet down')

    def test_address_absoute_link_sheet_n_workbook_token(self):
        cell_values = self.executor.get_cells([Cell('address', 7, 2), Cell('address', 7, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='ADDRESS absolute link sheet & workbook down')

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

    def test_column_no_args_token(self):
        cell_values = self.executor.get_cells([Cell('column', 0, 1), Cell('column', 0, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='Column token no args down')

    def test_column_cell_arg_token(self):
        cell_values = self.executor.get_cells([Cell('column', 1, 1), Cell('column', 0, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='Column token cell arg down')

    def test_column_matrix_arg_token(self):
        expected = self.executor.get_cells([Cell('column', 2, 1), Cell('column', 3, 1), Cell('column', 4, 1)])
        actual = self.executor.get_cells([Cell('column', 0, 4), Cell('column', 1, 4), Cell('column', 2, 4)])

        self.assertEqual([cell.value for cell in expected], [cell.value for cell in actual], msg='Column token matrix arg down')

    def test_column_array_arg_token(self):
        cell_values = self.executor.get_cells([Cell('column', 5, 1), Cell('column', 0, 5)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='Column token array arg down')

    def test_today_token(self):
        cell_values = self.executor.get_cells([Cell('today', 0, 2), Cell('today', 0, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='TODAY token down')

    def test_today_subtract_date_token(self):
        cell_values = self.executor.get_cells([Cell('today', 2, 2), Cell('today', 2, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value.days, msg='TODAY subtract date token down')

    def test_today_get_day_token(self):
        cell_values = self.executor.get_cells([Cell('today', 3, 2), Cell('today', 3, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='TODAY get day token down')

    def test_today_get_month_token(self):
        cell_values = self.executor.get_cells([Cell('today', 3, 2), Cell('today', 3, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='TODAY get month token down')

    def test_compare_dates(self):
        cell_values = self.executor.get_cells([Cell('today', 4, 2), Cell('today', 4, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='date compare down')

    def test_sumifs_one_condition(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('sumifs', 3, 0), Cell('sumifs', 4, 0)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIFS one condition down')

    def test_sumifs_several_conditions(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('sumifs', 3, 1), Cell('sumifs', 4, 1)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIFS several condition down')

    def test_sumifs_starts_with_letter(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('sumifs', 3, 2), Cell('sumifs', 4, 2)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIFS starts with letter down')

    def test_sumifs_with_boolean_fields(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('sumifs', 3, 3), Cell('sumifs', 4, 3)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIFS with boolean fields down')

    def test_sumifs_several_conditions_in_one_range(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('sumifs', 3, 4), Cell('sumifs', 4, 4)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIFS several conditions in one range down')

    def test_sumifs_more_than(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('sumifs', 3, 5), Cell('sumifs', 4, 5)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIFS more than down')

    def test_sumifs_less_than(self):
        cell_values = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cells([Cell('sumifs', 3, 6), Cell('sumifs', 4, 6)])

        self.assertEqual(cell_values[0].value, cell_values[1].value, msg='SUMIFS less than down')

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

    def test_index_merge_sets_values(self):
        cell_value = self.executor.get_cell(Cell('index', 0, 5))

        self.assertEqual(cell_value.value, [['12'], ['45'], ['78']])

    def test_roundup_with_digits_token(self):
        cell_values = self.executor.get_cells([Cell('roundup', 1, 0), Cell('roundup', 1, 1), Cell('roundup', 1, 2)])

        self.assertEqual(cell_values[0].value, 3.2, msg='ROUNDUP token with digits are OK')
        self.assertEqual(cell_values[1].value, 3.15, msg='ROUNDUP token with digits are OK')
        self.assertEqual(cell_values[2].value, 3.15, msg='ROUNDUP token with digits are OK')

    def test_roundup_with_separator_token(self):
        cell_values = self.executor.get_cells([Cell('roundup', 0, 3), Cell('roundup', 0, 4), Cell('roundup', 0, 5)])

        self.assertEqual(cell_values[0].value, 4, msg='ROUNDUP token with separator are OK')
        self.assertEqual(cell_values[1].value, 4, msg='ROUNDUP token with separator are OK')
        self.assertEqual(cell_values[2].value, 4, msg='ROUNDUP token with separator are OK')

    def test_roundup_without_digits_token(self):
        cell_values = self.executor.get_cells([Cell('roundup', 0, 0), Cell('roundup', 0, 1), Cell('roundup', 0, 2)])

        self.assertEqual(cell_values[0].value, 4, msg='ROUNDUP token without digits are OK')
        self.assertEqual(cell_values[1].value, 4, msg='ROUNDUP token without digits are OK')
        self.assertEqual(cell_values[2].value, 4, msg='ROUNDUP token without digits are OK')

    def test_rounddown_with_digits_token(self):
        cell_values = self.executor.get_cells(
            [Cell('rounddown', 1, 0), Cell('rounddown', 1, 1), Cell('rounddown', 1, 2)]
        )

        self.assertEqual(cell_values[0].value, 3.1, msg='ROUNDDOWN token with digits are OK')
        self.assertEqual(cell_values[1].value, 3.14, msg='ROUNDDOWN token with digits are OK')
        self.assertEqual(cell_values[2].value, 3.14, msg='ROUNDDOWN token with digits are OK')

    def test_rounddown_with_separator_token(self):
        cell_values = self.executor.get_cells(
            [Cell('rounddown', 0, 3), Cell('rounddown', 0, 4), Cell('rounddown', 0, 5)]
        )

        self.assertEqual(cell_values[0].value, 3, msg='ROUNDDOWN token with separator are OK')
        self.assertEqual(cell_values[1].value, 3, msg='ROUNDDOWN token with separator are OK')
        self.assertEqual(cell_values[2].value, 3, msg='ROUNDDOWN token with separator are OK')

    def test_rounddown_without_digits_token(self):
        cell_values = self.executor.get_cells(
            [Cell('rounddown', 0, 0), Cell('rounddown', 0, 1), Cell('rounddown', 0, 2)]
        )

        self.assertEqual(cell_values[0].value, 3, msg='ROUNDDOWN token without digits are OK')
        self.assertEqual(cell_values[1].value, 3, msg='ROUNDDOWN token without digits are OK')
        self.assertEqual(cell_values[2].value, 3, msg='ROUNDDOWN token without digits are OK')

    def test_left_operand_expression_percent_token(self):
        cell_values = self.executor.get_cells([
            Cell('leftOperandExpression', 0, 0), Cell('leftOperandExpression', 0, 1),
            Cell('leftOperandExpression', 0, 2), Cell('leftOperandExpression', 0, 3),
            Cell('leftOperandExpression', 0, 4)
        ])

        self.assertEqual(cell_values[0].value, 0.07, msg='Left operand expression percent operator are OK')
        self.assertEqual(cell_values[1].value, 0.14, msg='Left operand expression percent operator are OK')
        self.assertEqual(cell_values[2].value, 0.84, msg='Left operand expression percent operator are OK')
        self.assertEqual(cell_values[3].value, 0.0084, msg='Left operand expression percent operator are OK')
        self.assertEqual(cell_values[4].value, 0.0007, msg='Left operand expression percent operator are OK')

    def test_value_token_int(self):
        cell = self.executor.get_cell(Cell('value', 0, 0))

        self.assertEqual(cell.value, 123, msg='VALUE token int are OK')

    def test_value_token_float(self):
        cell = self.executor.get_cell(Cell('value', 0, 1))

        self.assertEqual(cell.value, 123.45, msg='VALUE token float are OK')

    def test_value_token_date(self):
        cell = self.executor.get_cell(Cell('value', 0, 2))

        self.assertEqual(cell.value, 45550, msg='VALUE token date are OK')

    def test_value_token_time(self):
        cell = self.executor.get_cell(Cell('value', 0, 3))

        self.assertEqual(cell.value, 0.5208333333333334, msg='VALUE token time are OK')

    def test_value_token_percent(self):
        cell = self.executor.get_cell(Cell('value', 0, 4))

        self.assertEqual(cell.value, 0.005, msg='VALUE token percent are OK')

    def test_value_token_error(self):
        cell = self.executor.get_cell(Cell('value', 0, 5))

        self.assertEqual(cell.value, '#VALUE!', msg='VALUE token error are OK')

    def test_value_token_datetime_expression(self):
        cell = self.executor.get_cell(Cell('value', 0, 6))

        self.assertEqual(cell.value, 45550, msg='VALUE token datetime expression are OK')

    def test_text_token(self):
        cell = self.executor.get_cell(Cell('text', 0, 0))

        self.assertEqual(cell.value, '1234567', msg='TEXT token are OK')

if __name__ == '__main__':
    unittest.main()
