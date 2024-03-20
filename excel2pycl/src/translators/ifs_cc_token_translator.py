from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IfsControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.utilities.helper import get_flatten_list


class IfsControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: IfsControlConstructionToken, excel: Excel, context: Context) -> str:
        flatten_list = get_flatten_list(token, excel, context)
        return context.set_sub_cell(token.in_cell, f'self._ifs({flatten_list})')
