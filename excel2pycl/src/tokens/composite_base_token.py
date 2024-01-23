from excel2pycl.src.cell import Cell
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens.base_token import BaseToken


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
        control_construction_flag = False
        for tokens in cls.get_token_sets():
            new_expression_part = []
            _expression = expression.copy()
            for token in tokens:
                if not len(_expression):
                    break
                elif token == _expression[0].__class__:
                    from excel2pycl.src.tokens import ControlConstructionCompositeBaseToken
                    """
                    If we encounter a control construction and have not encountered any of its token sets,
                    we are sure that the structure we are trying to parse contains an error
                    and further selection of tokens are not needed
                    """
                    control_construction_flag = cls in [
                        token[0] for token in ControlConstructionCompositeBaseToken.get_token_sets()
                    ]
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

        if control_construction_flag:
            raise E2PyclParserException(f'{cls.__name__} has an incorrect structure')

        return None, expression