from excel2pycl.src.cell import Cell
from excel2pycl.src.tokens import RegexpBaseToken, WhitespaceToken
from excel2pycl.src.tokens.base_token import BaseToken


class Lexer:
    TOKENS = RegexpBaseToken.subclasses()

    @classmethod
    def parse(cls, expression: str, in_cell: Cell) -> list[BaseToken]:
        tokens = []
        while expression:
            for token_class in cls.TOKENS:
                token, sub_expression = token_class.get(expression.lstrip(), in_cell)
                if token and token.__class__ is not WhitespaceToken:
                    tokens.append(token)
                    expression = sub_expression
                    break

        return tokens
