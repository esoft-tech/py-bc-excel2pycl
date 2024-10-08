from __future__ import annotations

from typing import List, Optional, Set, Dict, Union

from excel2pycl.src.handle_cell import handle_cell
from excel2pycl.src.utilities.abstract_excel_in_python_class import AbstractExcelInPython
from excel2pycl.src.cell import Cell
from excel2pycl.src.exceptions import E2PyclExecutorException
from excel2pycl.src.object_loader import load_module


class Executor:
    def __init__(self):
        """
        Simplified model of working with translated ExcelInPython objects.
        """
        self._cells_have_been_changed: bool = False
        self._executed_instance: Optional[AbstractExcelInPython] = None
        self._cells: Set[Cell] = set()
        self._titles: Dict[str, int] = {}
        self._sheets_size: List[Dict[str, int]] = []

    def set_executed_class(self, class_object: AbstractExcelInPython.__class__ = None,
                           class_file: str = None) -> Executor:
        """
        Sets the executing ExcelInPython object.

        Args:
            class_object (AbstractExcelInPython.__class__): ExcelInPython class object.
            class_file (str): The absolute or relative path to the file containing the ExcelInPython class.

        Returns:
            Executor.

        Raises:
            E2PyclExecutorException: An exception will be thrown when no one of the arguments are passed.
        """
        if class_object:
            self._executed_instance = class_object()
        elif class_file:
            self._executed_instance = load_module(class_file).ExcelInPython()
        else:
            raise E2PyclExecutorException('There is no data to get an instance of an excel in python object.')

        self._titles = self._executed_instance.get_titles()
        self._sheets_size = self._executed_instance.get_sheets_size()

        return self

    def get_executed_class(self):
        return self._executed_instance

    def set_cells(self, cells: List[Cell]) -> Executor:
        """
        Sets cell values for an ExcelInPython object.
        Replaces the existing values at the intersection.

        Args:
            cells (List[Cell]): Cells with values.

        Returns:
            Executor.
        """
        for cell in cells:
            handle_cell(cell, self._titles)

            sheet = cell.title
            row = cell.row + 1
            column = cell.column + 1

            self._sheets_size[sheet]['last_row'] = max(row, self._sheets_size[sheet]['last_row'])
            self._sheets_size[sheet]['last_column'] = max(column, self._sheets_size[sheet]['last_column'])

        self._cells = {*cells, *self._cells}
        self._cells_have_been_changed = True
        return self

    def _set_cells_to_executed_instance(self) -> Executor:
        """
        Writes the available cell values to an ExcelInPython object.

        Returns:
            Executor.
        """
        self._executed_instance.set_arguments([cell.to_dict() for cell in self._cells])
        self._cells_have_been_changed = False
        return self

    def get_cell(self, cell: Cell) -> Cell:
        """
        Calculates the value of the passed cell in the ExcelInPython
        object and writes it to the passed instance of the cell.
        Returns the cell with the recorded value.

        Args:
            cell (Cell): The cell with the address whose value we want to get.

        Returns:
            Cell.
        """
        if self._cells_have_been_changed:
            self._set_cells_to_executed_instance()

        handle_cell(cell, self._titles)
        cell.value = self._executed_instance.exec_function_in(cell.uid)

        return cell

    def get_cells(self, cells: List[Cell]) -> List[Cell]:
        """
        Calculates the value of the passed cells in the ExcelInPython
        object and writes them to the passed instance of the cells.
        Returns the cells list with the recorded values.

        Args:
            cells (List[Cell]): List of cells with the addresses whose values we want to get.

        Returns:
            List[Cell].
        """
        return [self.get_cell(cell) for cell in cells]

    def get_sheet(self, sheet: Union[int, str]) -> List[List[Cell]]:
        """
        Calculates the value of the passed worksheet cells in the ExcelInPython
        object and writes them to the passed instance of the cells.
        Returns the list of cell lists with the recorded values.

        Args:
            sheet (Union[int, str]): Index or name of the worksheet which cell values we want to get.

        Returns:
            List[List[Cell]].
        """
        if type(sheet) is str:
            sheet = self._titles[sheet]

        cells = []
        sheet_size = self._sheets_size[sheet]

        for row in range(sheet_size.get('last_row', 0)):
            row_cells = []
            for column in range(sheet_size.get('last_column', 0)):
                row_cells.append(self.get_cell(Cell(title=sheet, row=row, column=column)))
            cells.append(row_cells)

        return cells
