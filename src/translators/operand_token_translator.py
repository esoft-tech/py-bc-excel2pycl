from src.context import Context
from src.excel import Excel
from src.tokens import OperandToken
from src.translators.abstract_translator import AbstractTranslator


class OperandTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: OperandToken, excel: Excel, context: Context) -> str:
        from src.translators.cell_translator import CellTranslator

        if token.cell:
            return CellTranslator.translate(token.cell, excel, context)
        elif token.range[0]:
            start_cell, finish_cell = token.range
            range_ = excel.get_range(start_cell, finish_cell)
            range_code = '[' + ','.join([CellTranslator.translate(i, excel, context) for i in range_]) + ']'
            return context.set_sub_cell(token.cell, range_code)
        elif token.matrix[0]:
            start_cell, finish_cell = token.range
            matrix = excel.get_matrix(start_cell, finish_cell)
            matrix_cell_codes = '[' + ','.join(
                ['[' + ','.join([CellTranslator.translate(j, excel, context) for j in i]) + ']' for i in matrix]) + ']'
            return context.set_sub_cell(token.cell, matrix_cell_codes)
        elif token.literal:
            return token.literal
        elif token.control_construction:
            from src.translators.cc_token_translator import ControlConstructionTokenTranslator
            return ControlConstructionTokenTranslator.translate(token.control_construction, excel, context)
        else:
            raise Exception('Undefined token value')
