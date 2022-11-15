from excel2pycl.src.cell import Cell
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens import EntryPointToken


class AstBuilder:
    @classmethod
    def parse(cls, expression: list, in_cell: Cell):
        ast = EntryPointToken.get(expression, in_cell)
        if len(ast[0].value[1].value) != len(expression) - 1:
            raise E2PyclParserException('Unknown tokens composition', ast[1][0], expression)
        return ast[0]
