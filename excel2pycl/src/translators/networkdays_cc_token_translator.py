from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import NetworkDaysControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class NetworkDaysControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: NetworkDaysControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        date_start = ExpressionTokenTranslator.translate(token.date_start, excel, context)
        date_end = ExpressionTokenTranslator.translate(token.date_end, excel, context)

        additional = ExpressionTokenTranslator.translate(token.additional, excel, context) if token.additional else None

        return f'self._network_days({date_start}, {date_end}, {additional})'
