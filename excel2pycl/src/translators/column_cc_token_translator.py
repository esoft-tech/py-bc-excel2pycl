from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ColumnControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ColumnControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: ColumnControlConstructionToken, excel: Excel, context: Context) -> str:
        if token.cell:
            return str(token.cell.cell.column + 1)
        elif token.matrix:
            return str(list(set([str(cell.column + 1) for cell in token.matrix.matrix])))
        else:
            return token.in_cell.column + 1
