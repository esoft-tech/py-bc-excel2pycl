from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.exceptions.e2pycl_translation_exception import E2PyclTranslationException
from excel2pycl.src.tokens import XMatchControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class XMatchControlConstructionTokenTranslator(AbstractTranslator[XMatchControlConstructionToken]):
    @classmethod
    def translate(cls, token: XMatchControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import (
            MatrixOfCellIdentifiersTokenTranslator,
        )

        if token.match_mode is None:
            raise E2PyclTranslationException("token.math_mode is None", token, context)

        lookup_value, lookup_array, match_mode, search_mode = (
            ExpressionTokenTranslator.translate(token.lookup_value, excel, context),
            MatrixOfCellIdentifiersTokenTranslator.translate(token.lookup_array, excel, context),
            ExpressionTokenTranslator.translate(token.match_mode, excel, context),
            ExpressionTokenTranslator.translate(token.search_mode, excel, context) if token.search_mode else True,
        )

        return context.set_sub_cell(
            token.in_cell,
            f"self._xmatch({lookup_value}, {lookup_array}, {match_mode}, {search_mode})",
        )
