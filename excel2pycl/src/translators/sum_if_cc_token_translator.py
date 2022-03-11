from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import SumIfControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class SumIfControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: SumIfControlConstructionToken, excel: Excel, context: Context) -> str:
        if token.cell:
            start, finish = token.cell, token.cell
        elif token.range:
            start, finish = token.range.range
        elif token.matrix:
            start, finish = token.matrix.matrix
        else:
            raise ValueError('Invalid cells range')

        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import MatrixOfCellIdentifiersTokenTranslator
        range_ = MatrixOfCellIdentifiersTokenTranslator.translate_with_cells(start, finish, excel, context)

        start_cell_of_needed = token.first_cell_of_needed or start
        finish_cell_of_needed = excel.get_similar_second(start_cell_of_needed, start, finish)
        sum_range = MatrixOfCellIdentifiersTokenTranslator.translate_with_cells(start_cell_of_needed,
                                                                                finish_cell_of_needed, excel, context)

        from excel2pycl.src.translators.lambda_token_translator import LambdaTokenTranslator
        lambda_ = LambdaTokenTranslator.translate(token.lambda_, excel, context)

        return f'self._sum_if({range_},{lambda_}, {sum_range})'
