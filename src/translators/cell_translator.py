from src.cell import Cell
from src.context import Context
from src.excel import Excel
from src.translators.abstract_translator import AbstractTranslator


class CellTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, cell: Cell, excel: Excel, context: Context) -> str:
        from src.translators.entry_point_token_translator import EntryPointTokenTranslator

        if not cell.has_handled_identifiers():
            excel.fill_cell(cell)
        if not context.get_cell(cell):
            if type(cell.value) is str and cell.value.find('=') == 0:
                from src.ast_builder import AstBuilder
                from src.lexer import Lexer
                code = EntryPointTokenTranslator.translate(
                    AstBuilder.parse(Lexer.parse(cell.value, in_cell=cell), in_cell=cell), excel, context)
            else:
                code = cell.value
                if type(code) is str:
                    code = f"'{code}'"
            context.set_cell(cell, code)
        return context.get_cell(cell)
