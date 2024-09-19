from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import RoundUpControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class RoundUpControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: RoundUpControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        number = ExpressionTokenTranslator.translate(token.number, excel, context)

        num_digits = '0'
        if token.num_digits:
            num_digits = ExpressionTokenTranslator.translate(token.num_digits, excel, context)

        return f'self._roundup({number}, {num_digits})'
