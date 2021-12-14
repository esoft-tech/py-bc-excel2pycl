from src.context import Context
from src.excel import Excel
from src.tokens import MatrixOfCellIdentifiersToken
from src.translators.abstract_translator import AbstractTranslator


class MatrixOfCellIdentifiersTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: MatrixOfCellIdentifiersToken, excel: Excel, context: Context) -> str:
        from src.translators.cell_translator import CellTranslator

        start_cell, finish_cell = token.matrix
        matrix = excel.get_matrix(start_cell, finish_cell)
        matrix_cell_codes = '[' + ','.join(
            ['[' + ','.join([CellTranslator.translate(j, excel, context) for j in i]) + ']' for i in matrix]) + ']'

        return context.set_sub_cell(token.in_cell, matrix_cell_codes)
