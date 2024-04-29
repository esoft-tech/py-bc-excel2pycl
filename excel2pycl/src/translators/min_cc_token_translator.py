from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import MinControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.utilities.helper import get_flatten_list


class MinControlConstructionTokenTranslator(AbstractTranslator[MinControlConstructionToken]):
    @classmethod
    def translate(cls, token: MinControlConstructionToken, excel: Excel, context: Context) -> str:
        flatten_numeric_list = get_flatten_list(token, excel, context)
        return context.set_sub_cell(token.in_cell, f"self._min({flatten_numeric_list})")
