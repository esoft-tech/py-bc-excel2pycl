from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import MatrixOfCellIdentifiersExpressionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class MatrixOfCellIdentifiersExpressionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: MatrixOfCellIdentifiersExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import MatrixOfCellIdentifiersTokenTranslator

        left, right = token.operands

        list1 = MatrixOfCellIdentifiersTokenTranslator.translate(left, excel, context)
        list2 = MatrixOfCellIdentifiersTokenTranslator.translate(right, excel, context)

        return context.set_sub_cell(token.in_cell, f'self._sum_arrays(self._flatten_list({list1}), self._flatten_list({list2}))')
