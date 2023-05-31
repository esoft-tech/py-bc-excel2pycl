from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import MatchControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class MatchControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: MatchControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import MatrixOfCellIdentifiersTokenTranslator
        lookup_value, lookup_array, match_type \
            = ExpressionTokenTranslator.translate(token.lookup_value, excel, context), \
              MatrixOfCellIdentifiersTokenTranslator.translate(token.lookup_array, excel, context), \
              ExpressionTokenTranslator.translate(token.match_type, excel, context)

        return context.set_sub_cell(token.in_cell, f'self._match({lookup_value}, {lookup_array}, {match_type})')

