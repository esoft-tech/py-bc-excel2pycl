from src.cell import Cell
from src.tokens.base_token import BaseToken


class CompositeBaseToken(BaseToken):
    _TOKEN_SETS = []
    _PROCESSED = False

    @classmethod
    def add_token_set(cls, tokens: list):
        cls._TOKEN_SETS.append(tokens)

    @classmethod
    def get_token_sets(cls) -> list:
        return cls._TOKEN_SETS

    @classmethod
    def get(cls, expression: list, in_cell: Cell):
        for tokens in cls.get_token_sets():
            new_expression_part = []
            _expression = expression.copy()
            for token in tokens:
                if not len(_expression):
                    break
                elif token == _expression[0].__class__:
                    new_expression_part.append(_expression[0])
                    _expression = _expression[1:]
                elif token in CompositeBaseToken.subclasses():
                    new_token, _expression = token.get(_expression, in_cell)
                    if not new_token:
                        break
                    new_expression_part.append(new_token)
                else:
                    break

            if len(new_expression_part) == len(tokens) and len(new_expression_part):
                return cls(new_expression_part, in_cell), _expression

        return None, expression
