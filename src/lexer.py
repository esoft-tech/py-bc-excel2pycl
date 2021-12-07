from src.cell import Cell
from src.tokens import RegexpBaseToken, WhitespaceToken


class Lexer:
    TOKENS = RegexpBaseToken.subclasses()

    @classmethod
    def parse(cls, expression, in_cell: Cell):
        tokens = []
        while expression:
            for token_class in cls.TOKENS:
                token, sub_expression = token_class.get(expression, in_cell)
                if token and token.__class__ is not WhitespaceToken:
                    tokens.append(token)
                    expression = sub_expression
                    break

        return tokens
