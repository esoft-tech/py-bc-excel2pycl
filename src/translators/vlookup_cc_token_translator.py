from src.context import Context
from src.excel import Excel
from src.tokens import VlookupControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class VlookupControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: VlookupControlConstructionToken, excel: Excel, context: Context) -> str:
        pass
