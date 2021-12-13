from src.context import Context
from src.excel import Excel
from src.tokens import AverageControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class AverageControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: AverageControlConstructionToken, excel: Excel, context: Context) -> str:
        pass
