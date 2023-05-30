from __future__ import annotations

from typing import Optional

from excel2pycl.src.cell import Cell
from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.translators import CellTranslator


class Parser:
    def __init__(self):
        """
        Simplified model of interaction with the translator.
        """
        self._safety_check: bool = True
        self._translation: Optional[str] = None
        self._entrypoint_cell: Optional[Cell] = None
        self._entrypoint_cell_has_been_changed: bool = True
        self._excel_file_path: Optional[str] = None
        self._excel_file_path_has_been_changed: bool = True

    def enable_safety_check(self) -> Parser:
        """
        Enables the security check of the Excel file.
        If suspicious areas are found in the file during translation, the E2PyclSafetyException will be thrown.

        Returns:
            Parser.
        """
        self._safety_check = True
        return self

    def disable_safety_check(self) -> Parser:
        """
        Disables the security check of the Excel file.

        Returns:
            Parser.
        """
        self._safety_check = False
        return self

    def set_excel_file_path(self, excel_file_path: str) -> Parser:
        """
        Sets the value of the file path for translation.

        Args:
            excel_file_path (str): The path to the file on the disk.

        Returns:
            Parser.
        """
        self._excel_file_path = excel_file_path
        self._excel_file_path_has_been_changed = True
        return self

    def set_entrypoint_cell(self, cell: Cell) -> Parser:
        """
        Sets the cell that serves as the entry point. Its dependencies will be parsed from this cell.

        Args:
            cell (Cell): Cell that is the entrypoint.

        Returns:
            Parser.
        """
        self._entrypoint_cell = cell
        return self

    def _translate(self) -> Parser:
        """
        Translates the Excel file to python code.

        Returns:
            Parser.

        Raises:
            E2PyclParserException: Throws exception in the absence of required fields or problems with parsing.
            E2PyclSafetyException: If security check is enabled and suspicious fragments are found,
                an exception will be thrown.
        """
        if not self._excel_file_path_has_been_changed and not self._entrypoint_cell_has_been_changed:
            return self

        if not self._excel_file_path:
            raise E2PyclParserException('The file path is not set.')

        excel = Excel.parse(self._excel_file_path)
        if self._safety_check:
            excel.is_safe()

        context = Context()
        context._titles = excel.get_titles()

        if self._entrypoint_cell:
            CellTranslator.translate(self._entrypoint_cell, excel, context)
        else:
            CellTranslator.translate_file(excel, context)

        self._translation = context.build_class()

        self._excel_file_path_has_been_changed = False
        self._entrypoint_cell_has_been_changed = False

        return self

    def get_translation(self) -> str:
        """
        Returns:
            str: The python class code resulting from the translation.

        Raises:
            E2PyclParserException: Throws exception in the absence of required fields or problems with parsing.
            E2PyclSafetyException: If security check is enabled and suspicious fragments are found,
                an exception will be thrown.
        """
        return self._translate()._translation

    def write_translation(self, file_path: str) -> Parser:
        """
        Writes the python class code resulting from the translation to a file.

        Args:
            file_path (str):  The file where to write the python class code resulting from the translation.

        Returns:
            Parser.

        Raises:
            E2PyclParserException: Throws exception in the absence of required fields or problems with parsing.
            E2PyclSafetyException: If security check is enabled and suspicious fragments are found,
                an exception will be thrown.
        """
        self._translate()

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(self._translation)

        return self
