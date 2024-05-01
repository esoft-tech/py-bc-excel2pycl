from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import CountControlConstructionToken
from excel2pycl.src.translators import CellTranslator
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import (
    MatrixOfCellIdentifiersTokenTranslator,
)
from excel2pycl.src.utilities.helper import get_flatten_list


class CountControlConstructionTokenTranslator(AbstractTranslator[CountControlConstructionToken]):
    @classmethod
    def translate(cls, token: CountControlConstructionToken, excel: Excel, context: Context) -> str:
        cls.check_token_type(token, CountControlConstructionToken)
        args = get_flatten_list(token, excel, context)
        matrices = (
            ", ".join(
                MatrixOfCellIdentifiersTokenTranslator.translate(matrix, excel, context) for matrix in token.matrices
            )
            or "[]"
        )
        arg_cells = ", ".join(CellTranslator.translate(arg.cell, excel, context) for arg in token.arg_cells)
        return f"self._count({matrices}, {args}, [{arg_cells}])"
