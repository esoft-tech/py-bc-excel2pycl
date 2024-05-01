from excel2pycl.src.cell import Cell
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens.base_token import BaseToken


class UndefinedToken(BaseToken):
    @classmethod
    def get(cls, expression: str | list, in_cell: Cell) -> tuple["BaseToken | None", list | str]:
        raise E2PyclParserException("Undefined token", expression, in_cell)
