from excel2pycl.src.cell import Cell
from excel2pycl.src.tokens import EntryPointToken
from excel2pycl.src.tokens.base_token import BaseToken


class AstBuilder:
    @classmethod
    def parse(cls, expression: list, in_cell: Cell) -> BaseToken | None:
        return EntryPointToken.get(expression, in_cell)[0]
