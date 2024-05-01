from typing import cast

from openpyxl.utils import column_index_from_string

from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ColumnControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ColumnControlConstructionTokenTranslator(AbstractTranslator[ColumnControlConstructionToken]):
    @classmethod
    def translate(cls, token: ColumnControlConstructionToken, excel: Excel, context: Context) -> str:
        cls.check_token_type(token, ColumnControlConstructionToken)
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import (
            MatrixOfCellIdentifiersTokenTranslator,
        )

        if token.cell:
            return str(column_index_from_string(str(token.cell.cell.column)))

        if token.matrix:
            # Mutates matrix, inplace literal cols with digital
            MatrixOfCellIdentifiersTokenTranslator.translate(token.matrix, excel, context)

            if token.matrix.matrix[0].column == token.matrix.matrix[-1].column:
                return str(cast(int, token.matrix.matrix[0].column) + 1)

            for i in range(cast(int, token.matrix.matrix[0].column) + 1, cast(int, token.matrix.matrix[-1].column) + 2):
                # The only way to set multiple cells while parsing single token
                context.set_cell(token.in_cell, str(i))
                token.in_cell.column += 1  # type: ignore [operator]
            token.in_cell.column = cast(int, token.matrix.matrix[0].column) + 1
            return context.set_sub_cell(token.in_cell, str(token.in_cell.column))

        return str(cast(int, token.in_cell.column) + 1)
