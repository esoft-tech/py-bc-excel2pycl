from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import EDateControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class EDateControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: EDateControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        start_date = ExpressionTokenTranslator.translate(token.start_date, excel, context)
        months = ExpressionTokenTranslator.translate(token.months, excel, context)

        return f'self._edate({start_date}, {months})'
