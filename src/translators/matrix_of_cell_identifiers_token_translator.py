from src.cell import Cell
from src.context import Context
from src.excel import Excel
from src.tokens import MatrixOfCellIdentifiersToken
from src.translators.abstract_translator import AbstractTranslator


class MatrixOfCellIdentifiersTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: MatrixOfCellIdentifiersToken, excel: Excel, context: Context) -> str:
        start_cell, finish_cell = token.matrix
        return cls.translate_with_cells(start_cell, finish_cell, excel, context)

    @classmethod
    def translate_with_cells(cls, start_cell: Cell, finish_cell: Cell, excel: Excel, context: Context) -> str:
        from src.translators.cell_translator import CellTranslator

        matrix = excel.get_matrix(start_cell, finish_cell)
        matrix_cell_codes = '[' + ','.join(
            ['[' + ','.join([CellTranslator.translate(j, excel, context) for j in i]) + ']' for i in matrix]) + ']'

        return context.set_sub_cell(start_cell, matrix_cell_codes)
