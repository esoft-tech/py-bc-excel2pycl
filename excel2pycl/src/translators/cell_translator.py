from excel2pycl.src.cell import Cell
from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.exceptions import E2PyclTranslationException
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class CellTranslator(AbstractTranslator[Cell]):
    @classmethod
    def _set_cell_to_context(cls, cell: Cell, excel: Excel, context: Context) -> tuple[Cell, Excel, Context]:
        """
        Translates the passed cell from the Excel object and all nested cells.

        Args:
            cell (Cell): The cell that will be translated.
            excel (Excel): The object of the Excel file to be translated.
            context (Context): The translation context where the translation maps will be stored.

        Returns:
            (Cell, Excel, Context).
        """
        from excel2pycl.src.translators.entry_point_token_translator import EntryPointTokenTranslator

        if not cell.has_handled_identifiers():
            excel.fill_cell(cell)
        if not context.get_cell(cell):
            if isinstance(cell.value, str) and cell.value.find("=") == 0:
                from excel2pycl.src.ast_builder import AstBuilder
                from excel2pycl.src.lexer import Lexer

                lexer = Lexer.parse(cell.value, in_cell=cell)
                ast = AstBuilder.parse(lexer, in_cell=cell)
                code = EntryPointTokenTranslator.translate(ast, excel, context)
            else:
                code = repr(cell.value) if cell.value is not None else "self.EmptyCell()"
            context.set_cell(cell, code)
        return cell, excel, context

    @classmethod
    def translate(cls, cell: Cell, excel: Excel, context: Context) -> str:
        """
        Translates the passed cell from the Excel object and all nested cells.

        Args:
            cell (Cell): The cell that will be translated.
            excel (Excel): The object of the Excel file to be translated.
            context (Context): The translation context where the translation maps will be stored.

        Returns:
            str: Translation of the passed cell.
        """
        cell, excel, context = cls._set_cell_to_context(cell, excel, context)
        translated_cell = context.get_cell(cell)

        if translated_cell is None:
            raise E2PyclTranslationException("Cannot translate a cell", cell, context)

        return translated_cell

    @classmethod
    def translate_file(cls, excel: Excel, context: Context) -> None:
        """
        Translates all cells from an Excel object.

        Args:
            excel (Excel): The object of the Excel file to be translated.
            context (Context): The translation context where the translation maps will be stored.
        """
        for cell in excel.get_cells():
            cell, excel, context = cls._set_cell_to_context(cell, excel, context)
