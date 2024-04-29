import datetime
import os
import unittest
import uuid

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from excel2pycl import Cell, Executor, Parser


def create_test_table(filename: str) -> None:
    wb = Workbook()
    ws = wb.active

    tab = Table(displayName="base", ref="A1:C5")

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=True,
    )
    tab.tableStyleInfo = style

    ws.add_table(tab)

    data = [
        [datetime.datetime(2023, 4, 1), datetime.datetime(2023, 5, 31), "=NETWORKDAYS(A1,B1)"],
        [datetime.datetime(2023, 5, 31), datetime.datetime(2023, 4, 1), "=NETWORKDAYS(A2,B2)"],
        [datetime.datetime(2023, 5, 1), datetime.datetime(2023, 5, 8), ""],
        [40, "ewewwewe", "=NETWORKDAYS(A1,B1,A3:B4)"],
        ["", "", "=NETWORKDAYS(A4,B4)"],
    ]

    for row in data:
        ws.append(row)

    wb.save(filename)


class TestNetworkDaysCcToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f"tests/{uuid.uuid4()}.py"
        create_test_table("tests/networks_days.xlsx")
        cls.parser = (
            Parser()
            .set_excel_file_path("tests/networks_days.xlsx")
            .enable_safety_check()
            .write_translation(cls.translation_file_path)
        )

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)
        os.remove("tests/networks_days.xlsx")

    def test_two_args(self):
        excepted_cell_value = 43
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 0))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_three_args(self):
        excepted_cell_value = 41
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 3))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_incorrect_interval(self):
        excepted_cell_value = "#VALUE!"
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 4))

        self.assertEqual(cell_value.value, excepted_cell_value)

    def test_date_start_bigger_date_end(self):
        excepted_cell_value = -43
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 1))

        self.assertEqual(cell_value.value, excepted_cell_value)
