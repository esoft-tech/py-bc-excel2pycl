from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import TodayControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class TodayControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: TodayControlConstructionToken, excel: Excel, context: Context) -> str:
        return f'self._today()'
