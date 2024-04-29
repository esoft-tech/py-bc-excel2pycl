from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import DateControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class DateControlConstructionTokenTranslator(AbstractTranslator[DateControlConstructionToken]):
    @classmethod
    def translate(cls, token: DateControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        year = ExpressionTokenTranslator.translate(token.year, excel, context)
        month = ExpressionTokenTranslator.translate(token.month, excel, context)
        day = ExpressionTokenTranslator.translate(token.day, excel, context)

        return f"self._date({year}, {month}, {day})"
