from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IterableRangeOfConditionsWithExpressionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IterableRangeOfConditionsWithExpressionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: IterableRangeOfConditionsWithExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.range_of_conditions_with_expression_token_translator import \
            RangeOfConditionsWithExpressionTokenTranslator
        return context.set_sub_cell(
            token.in_cell, '[' + ', '.join(
                [f'*{RangeOfConditionsWithExpressionTokenTranslator.translate(i, excel, context)}' for i in
                    token.range_of_conditions_with_expressions]
                ) + ']'
            )