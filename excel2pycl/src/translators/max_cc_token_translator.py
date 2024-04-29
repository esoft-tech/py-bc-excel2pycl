from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import MaxControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.utilities.helper import get_flatten_numeric_list


class MaxControlConstructionTokenTranslator(AbstractTranslator[MaxControlConstructionToken]):
    @classmethod
    def translate(cls, token: MaxControlConstructionToken, excel: Excel, context: Context) -> str:
        flatten_numeric_list = get_flatten_numeric_list(token, excel, context)
        return context.set_sub_cell(token.in_cell, f"self._max({flatten_numeric_list})")
