from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import MonthControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class MonthControlConstructionTokenTranslator(AbstractTranslator[MonthControlConstructionToken]):
    @classmethod
    def translate(cls, token: MonthControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        date = ExpressionTokenTranslator.translate(token.date, excel, context)

        return f"self._month({date})"
