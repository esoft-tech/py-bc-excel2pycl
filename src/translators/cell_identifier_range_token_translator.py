from src.context import Context
from src.excel import Excel
from src.tokens import CellIdentifierRangeToken
from src.translators.abstract_translator import AbstractTranslator


class CellIdentifierRangeTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: CellIdentifierRangeToken, excel: Excel, context: Context) -> str:
        from src.translators.cell_translator import CellTranslator

        start_cell, finish_cell = token.range
        range_ = excel.get_range(start_cell, finish_cell)
        range_code = '[' + ','.join([CellTranslator.translate(i, excel, context) for i in range_]) + ']'

        return context.set_sub_cell(token.in_cell, range_code)
