from src.context import Context
from src.excel import Excel
from src.tokens import VlookupControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class VlookupControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: VlookupControlConstructionToken, excel: Excel, context: Context) -> str:
        from src.translators import ExpressionTokenTranslator
        from src.translators.matrix_of_cell_identifiers_token_translator import MatrixOfCellIdentifiersTokenTranslator
        lookup_value, matrix, column_number, range_lookup \
            = ExpressionTokenTranslator.translate(token.lookup_value, excel, context), \
              MatrixOfCellIdentifiersTokenTranslator.translate(token.matrix, excel, context), \
              ExpressionTokenTranslator.translate(token.column_number, excel, context), \
              ExpressionTokenTranslator.translate(token.range_lookup, excel, context) if token.range_lookup else 'False'

        return context.set_sub_cell(token.in_cell, f'self._vlookup({lookup_value}, {matrix}, {column_number}, {range_lookup})')

