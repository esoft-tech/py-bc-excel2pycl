from typing import cast

from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import CellIdentifierRangeToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class CellIdentifierRangeTokenTranslator(AbstractTranslator[CellIdentifierRangeToken]):
    @classmethod
    def translate(cls, token: CellIdentifierRangeToken, excel: Excel, context: Context) -> str:
        cls.check_token_type(token, CellIdentifierRangeToken)
        token = cast(CellIdentifierRangeToken, token)

        from excel2pycl.src.translators.cell_translator import CellTranslator

        start_cell, finish_cell = token.range
        range_ = excel.get_range(start_cell, finish_cell)
        range_code = "[" + ",".join([CellTranslator.translate(i, excel, context) for i in range_]) + "]"

        return context.set_sub_cell(token.in_cell, range_code)
