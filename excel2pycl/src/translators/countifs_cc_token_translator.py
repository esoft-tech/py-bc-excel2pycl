from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import CountIfsControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.translators.lambda_token_translator import LambdaTokenTranslator


class CountIfsControlConstructionTokenTranslator(AbstractTranslator[CountIfsControlConstructionToken]):
    @classmethod
    def translate(cls, token: CountIfsControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.iterable_range_of_cell_identifier_with_condition_token_translator import (
            IterableRangeOfCellIdentifierWithConditionTokenTranslator,
        )
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import (
            MatrixOfCellIdentifiersTokenTranslator,
        )

        count_range = MatrixOfCellIdentifiersTokenTranslator.translate(token.count_range, excel, context)
        count_condition = LambdaTokenTranslator.translate(token.count_condition, excel, context)
        conditions = (
            IterableRangeOfCellIdentifierWithConditionTokenTranslator.translate(
                token.conditions,
                excel,
                context,
            )
            if token.conditions
            else []
        )
        return f"self._countifs({count_range}, {count_condition}, *{conditions})"
