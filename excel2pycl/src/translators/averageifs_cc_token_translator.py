from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import AverageIfsControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class AverageIfsControlConstructionTokenTranslator(AbstractTranslator[AverageIfsControlConstructionToken]):
    @classmethod
    def translate(cls, token: AverageIfsControlConstructionToken, excel: Excel, context: Context) -> str:
        cls.check_token_type(token, AverageIfsControlConstructionToken)

        from excel2pycl.src.translators.iterable_range_of_cell_identifier_with_condition_token_translator import (
            IterableRangeOfCellIdentifierWithConditionTokenTranslator,
        )
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import (
            MatrixOfCellIdentifiersTokenTranslator,
        )

        average_range = MatrixOfCellIdentifiersTokenTranslator.translate(token.average_range, excel, context)
        conditions = IterableRangeOfCellIdentifierWithConditionTokenTranslator.translate(
            token.conditions,
            excel,
            context,
        )

        return context.set_sub_cell(token.in_cell, f"self._averageifs({average_range}, *{conditions})")
