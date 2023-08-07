from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import PatternToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class PatternTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: PatternToken, excel: Excel, context: Context) -> str:
        return f'self._regexp({token.value[0]})'
