from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import SumIfsControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class SumIfsControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: SumIfsControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import \
            MatrixOfCellIdentifiersTokenTranslator
        from excel2pycl.src.translators.iterable_range_of_cell_identifier_with_condition_token_translator import \
            IterableRangeOfCellIdentifierWithConditionTokenTranslator
        sum_range = MatrixOfCellIdentifiersTokenTranslator.translate(token.sum_range, excel, context)
        conditions = IterableRangeOfCellIdentifierWithConditionTokenTranslator.translate(
            token.conditions, excel, context
        )

        return context.set_sub_cell(token.in_cell, f'self._sumifs({sum_range}, *{conditions})')
