from src.tokens.undefined_token import UndefinedToken
from src.tokens.base_token import BaseToken


class CompositeBaseToken(BaseToken):
    _TOKEN_SETS = []
    _PROCESSED = False

    @classmethod
    def add_token_set(cls, tokens: list):
        cls._TOKEN_SETS.append(tokens)

    @classmethod
    def subclasses(cls) -> list:
        from src.tokens.recursive_composite_base_token import RecursiveCompositeBaseToken

        subclasses = set(cls.__subclasses__() + RecursiveCompositeBaseToken.subclasses())
        subclasses.difference_update({RecursiveCompositeBaseToken})

        return list(subclasses) + [UndefinedToken]

    @classmethod
    def get_token_sets(cls) -> list:
        return cls._TOKEN_SETS

    @classmethod
    def get(cls, expression: list):
        for tokens in cls.get_token_sets():
            new_expression_part = []
            _expression = expression.copy()
            for token in tokens:
                if not len(_expression):
                    break
                elif token == _expression[0].__class__:
                    new_expression_part.append(_expression[0])
                    _expression = _expression[1:]
                else:
                    new_token, _expression = token.get(_expression)
                    if not new_token:
                        break
                    new_expression_part.append(new_token)

            if len(new_expression_part) == len(tokens):
                return cls(new_expression_part), _expression

        return None, expression
