from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens import OperandToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class OperandTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: OperandToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.cell_translator import CellTranslator

        if token.cell:
            return CellTranslator.translate(token.cell, excel, context)
        elif token.range:
            from excel2pycl.src.translators.cell_identifier_range_token_translator import \
                CellIdentifierRangeTokenTranslator
            return CellIdentifierRangeTokenTranslator.translate(token.range, excel, context)
        elif token.matrix:
            from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import \
                MatrixOfCellIdentifiersTokenTranslator
            return MatrixOfCellIdentifiersTokenTranslator.translate(token.matrix, excel, context)
        elif token.literal:
            return token.literal
        elif token.control_construction:
            from excel2pycl.src.translators.cc_token_translator import ControlConstructionTokenTranslator
            return ControlConstructionTokenTranslator.translate(token.control_construction, excel, context)
        else:
            raise E2PyclParserException('Undefined token value')
