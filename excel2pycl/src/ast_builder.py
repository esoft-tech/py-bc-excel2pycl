from excel2pycl.src.cell import Cell
from excel2pycl.src.tokens import EntryPointToken


class AstBuilder:
    @classmethod
    def parse(cls, expression: list, in_cell: Cell):
        return EntryPointToken.get(expression, in_cell)[0]
