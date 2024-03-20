from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IfsControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IfsControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: IfsControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.iterable_range_of_conditions_with_expression_token_translator import \
            IterableRangeOfConditionsWithExpressionTokenTranslator

        items = IterableRangeOfConditionsWithExpressionTokenTranslator.translate(
            token.items, excel, context
        ) if token.items else []

        return f'self._ifs({items})'
