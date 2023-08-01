from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import AddressControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class AddressControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, subject: AddressControlConstructionToken, excel: Excel, context: Context) -> str:
        pass
