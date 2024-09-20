from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import RoundDownControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class RoundDownControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: RoundDownControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        number = ExpressionTokenTranslator.translate(token.number, excel, context)

        num_digits = '0'
        if token.num_digits:
            num_digits = ExpressionTokenTranslator.translate(token.num_digits, excel, context)

        return f'self._rounddown({number}, {num_digits})'
