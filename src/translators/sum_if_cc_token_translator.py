from src.context import Context
from src.excel import Excel
from src.tokens import SumIfControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class SumIfControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: SumIfControlConstructionToken, excel: Excel, context: Context) -> str:
        pass
