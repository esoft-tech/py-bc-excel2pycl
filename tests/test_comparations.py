import os
import unittest
import uuid
from datetime import date, datetime
from typing import cast

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from excel2pycl import Cell, Executor, Parser


def create_test_table(file_name: str) -> None:
    wb = Workbook()
    ws: Worksheet = cast(Worksheet, wb.active)

    data = [
        [date(2024, 1, 1), datetime(2024, 1, 1, 0, 0, 0), "=A1=B1"],  # True
        [date(2024, 1, 1), datetime(2024, 1, 1, 1, 10, 10), "=A2=B2"],  # False
        [date(2024, 1, 1), datetime(2024, 1, 1, 0, 0, 0), "=A3>B3"],  # False
        [date(2024, 1, 1), datetime(2024, 1, 1, 0, 0, 0), "=A4<B4"],  # False
        ["", datetime(2024, 1, 1, 0, 0, 0), "=A5<B5"],  # True
        ["", datetime(2024, 1, 1, 0, 0, 0), "=A6<=B6"],  # True
        ["", datetime(2024, 1, 1, 0, 0, 0), "=A7>=B7"],  # False
    ]

    for row in data:
        ws.append(row)

    tab = Table(displayName="base", ref="A1:C10")

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

    wb.save(file_name)


class TestComparations(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.translation_file_path = f"tests/{uuid.uuid4()}.py"  # type: ignore [attr-defined]
        create_test_table("tests/comparations.xlsx")
        cls.parser = (  # type: ignore [attr-defined]
            Parser()
            .set_excel_file_path("tests/comparations.xlsx")
            .enable_safety_check()
            .write_translation(cls.translation_file_path)  # type: ignore [attr-defined]
        )

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)  # type: ignore [attr-defined]
        os.remove("tests/comparations.xlsx")

    def test_date_eq(self) -> None:
        excepted_cell_value = True
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 0))  # type: ignore [attr-defined]

        self.assertEqual(cell_value.value, excepted_cell_value, msg="test_date_eq")

    def test_date_not_eq(self) -> None:
        excepted_cell_value = False
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 1))  # type: ignore [attr-defined]

        self.assertEqual(cell_value.value, excepted_cell_value, msg="test_date_not_eq")

    def test_date_not_gt(self) -> None:
        excepted_cell_value = False
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 2))  # type: ignore [attr-defined]

        self.assertEqual(cell_value.value, excepted_cell_value, msg="test_date_not_gt")

    def test_date_not_lt(self) -> None:
        excepted_cell_value = False
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 3))  # type: ignore [attr-defined]

        self.assertEqual(cell_value.value, excepted_cell_value, msg="test_date_not_lt")

    def test_empty_cell_lt(self) -> None:
        excepted_cell_value = True
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 4))  # type: ignore [attr-defined]

        self.assertEqual(cell_value.value, excepted_cell_value, msg="test_empty_cell_lt")

    def test_empty_cell_le(self) -> None:
        excepted_cell_value = True
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 5))  # type: ignore [attr-defined]

        self.assertEqual(cell_value.value, excepted_cell_value, msg="test_empty_cell_le")

    def test_empty_cell_ge(self) -> None:
        excepted_cell_value = False
        cell_value = Executor().set_executed_class(class_file=self.translation_file_path).get_cell(Cell(0, 2, 6))  # type: ignore [attr-defined]

        self.assertEqual(cell_value.value, excepted_cell_value, msg="test_empty_cell_ge")
