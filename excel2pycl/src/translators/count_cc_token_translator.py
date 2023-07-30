from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import CountControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.utilities.helper import get_flatten_numeric_list


class CountControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: CountControlConstructionToken, excel: Excel, context: Context) -> str:
        flatten_numeric_list = get_flatten_numeric_list(token, excel, context)

        return f'self._count({flatten_numeric_list})'
