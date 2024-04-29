import os
import unittest
import uuid

from excel2pycl import Cell, Executor, Parser


class TestAverageIfsCcToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f"tests/{uuid.uuid4()}.py"
        cls.parser = (
            Parser()
            .set_excel_file_path("tests/Excel2PyCl.xlsx")
            .enable_safety_check()
            .write_translation(cls.translation_file_path)
        )

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)

    def test_only_standard_int_and_gt_condition(self):
        excepted_cell_value = 5
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 0))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_with_string_eq_in_condition(self):
        excepted_cell_value = 1
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 1))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_where_has_unreachable_string_condition(self):
        excepted_cell_value = "#DIV/0!"
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 2))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_where_average_range_has_empty_cell(self):
        excepted_cell_value = "#DIV/0!"
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 3))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_average_rage_with_bool_value(self):
        excepted_cell_value = 1
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 4))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_average_by_matrix(self):
        excepted_cell_value = 3
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 5))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_with_string_date_eq_condition(self):
        excepted_cell_value = 5
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 6))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_with_string_in_average_range(self):
        excepted_cell_value = "#DIV0!"
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 7, 7))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_with_different_range_sizes(self):
        executor = Executor().set_executed_class(class_file=self.translation_file_path)

        with self.assertRaises(executor.get_executed_class().ExcelInPythonException):
            executor.get_cell(Cell(0, 7, 8))


if __name__ == "__main__":
    unittest.main()
