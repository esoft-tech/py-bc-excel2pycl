from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IndexControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IndexControlConstructionTokenTranslator(AbstractTranslator[IndexControlConstructionToken]):
    @classmethod
    def translate(cls, token: IndexControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import (
            MatrixOfCellIdentifiersTokenTranslator,
        )

        matrix_list = context.set_sub_cell(
            token.in_cell,
            ",".join([MatrixOfCellIdentifiersTokenTranslator.translate(i, excel, context) for i in token.matrix_list]),
        )

        row_number, column_number, area_number = (
            ExpressionTokenTranslator.translate(token.row_number, excel, context),
            ExpressionTokenTranslator.translate(token.column_number, excel, context) if token.column_number else None,
            ExpressionTokenTranslator.translate(token.area_number, excel, context) if token.area_number else 1,
        )

        return context.set_sub_cell(
            token.in_cell,
            f"self._index({matrix_list}, {row_number}, {column_number}, {area_number})",
        )
