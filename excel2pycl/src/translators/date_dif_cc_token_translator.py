from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import DateDifControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class DateDifControlConstructionTokenTranslator(AbstractTranslator[DateDifControlConstructionToken]):
    @classmethod
    def translate(cls, token: DateDifControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        date_start = ExpressionTokenTranslator.translate(token.date_start, excel, context)
        date_end = ExpressionTokenTranslator.translate(token.date_end, excel, context)
        mode = ExpressionTokenTranslator.translate(token.mode, excel, context)

        return f"self._datedif({date_start}, {date_end}, {mode})"
