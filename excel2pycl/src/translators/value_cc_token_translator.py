from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ValueControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ValueControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: ValueControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        value = ExpressionTokenTranslator.translate(token.expression, excel, context)

        return f'self._value(str({value}))'
