from src.context import Context
from src.excel import Excel
from src.tokens import SumControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class SumControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: SumControlConstructionToken, excel: Excel, context: Context) -> str:
        pass
