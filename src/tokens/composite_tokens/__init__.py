from src.tokens.composite_base_token import CompositeBaseToken
from src.tokens.recursive_composite_base_token import RecursiveCompositeBaseToken, CLS
from src.tokens.regexp_tokens import *


class LogicalOperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[EqOperatorToken], [NotEqOperatorToken], [GtOperatorToken], [GtOrEqualOperatorToken],
                   [LtOperatorToken], [LtOrEqualOperatorToken]]

    @property
    def operator(self):
        return self.value[0]


class OneOperandArithmeticOperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[PlusOperatorToken], [MinusOperatorToken]]

    @property
    def operator(self):
        return self.value[0]


class ArithmeticOperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[OneOperandArithmeticOperatorToken], [MultiplicationOperatorToken], [DivOperatorToken]]

    @property
    def operator(self):
        return self.value[0].operator if self.value[0].__class__ is OneOperandArithmeticOperatorToken else self.value[0]


class OperandToken(CompositeBaseToken):
    _TOKEN_SETS = [[LiteralToken], [CellIdentifierToken], [CellIdentifierRangeToken], [MatrixOfCellIdentifiersToken]]

    @property
    def cell(self) -> Cell:
        return self.value[0].cell if self.value[0].__class__ == CellIdentifierToken else None

    @property
    def range(self) -> (Cell, Cell, ):
        return self.value[0].range if self.value[0].__class__ == CellIdentifierRangeToken else (None, None)

    @property
    def matrix(self) -> (Cell, Cell, ):
        return self.value[0].matrix if self.value[0].__class__ == MatrixOfCellIdentifiersToken else (None, None)

    @property
    def literal(self) -> int or float or str or bool:
        return self.value[0].value if self.value[0].__class__ == LiteralToken else None

    @property
    def control_construction(self):
        return self.value[0] if self.value[0].__class__ == ControlConstructionToken else None


class OperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[ArithmeticOperatorToken], [LogicalOperatorToken]]

    @property
    def operator(self):
        return self.value[0].operator


class ExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [[OperandToken, OperatorToken, CLS], [OneOperandArithmeticOperatorToken, CLS],
                   [BracketStartToken, CLS, BracketFinishToken], [OperandToken]]

    @property
    def in_brackets(self) -> bool:
        return self.value[0].__class__ in [OneOperandArithmeticOperatorToken, BracketStartToken]

    @property
    def operator(self):
        return self.value[0].operator if self.value[0].__class__ is OperatorToken else self.value[1].operator if len(
            self.value) > 1 and self.value[1].__class__ is OperatorToken else None

    @property
    def left_operand(self) -> OperandToken:
        return self.value[0] if self.value[0].__class__ is OperandToken else None

    @property
    def right_operand(self):
        return self.value[2] if len(self.value) == 3 and self.value[2].__class__ is self.__class__ else self.value[1] if len(self.value) >= 2 and self.value[1].__class__ is self.__class__ else None


class IterableExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [[ExpressionToken, SeparatorToken, CLS], [ExpressionToken]]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, *kwargs)
        self._expressions = []

    @property
    def expressions(self):
        if not self._expressions:
            self._expressions = [self.value[0]] + self.value[2].expressions if len(self.value) == 3 else [self.value[0]]
        return self._expressions


class LambdaToken(CompositeBaseToken):
    _TOKEN_SETS = [[LiteralToken, AndLambdaToken, ExpressionToken], [ExpressionToken]]


class IfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, SeparatorToken,
                    ExpressionToken, BracketFinishToken]]

    @property
    def condition(self) -> ExpressionToken:
        return self.value[2]

    @property
    def when_true(self) -> ExpressionToken:
        return self.value[4]

    @property
    def when_false(self) -> ExpressionToken:
        return self.value[6]


class SumControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class SumIfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumIfKeywordToken, BracketStartToken, CellIdentifierRangeToken, SeparatorToken, LambdaToken,
                    BracketFinishToken],
                   [SumIfKeywordToken, BracketStartToken, CellIdentifierRangeToken, SeparatorToken, LambdaToken,
                    SeparatorToken, CellIdentifierRangeToken, BracketFinishToken]]


class VlookupControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [
        [VlookupKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         SeparatorToken, ExpressionToken, BracketFinishToken]]


class AverageControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[AverageKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class ControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfControlConstructionToken], [SumControlConstructionToken], [SumIfControlConstructionToken],
                   [VlookupControlConstructionToken], [AverageControlConstructionToken]]

    @property
    def control_construction(
            self) -> IfControlConstructionToken or SumControlConstructionToken or SumIfControlConstructionToken or VlookupControlConstructionToken or AverageControlConstructionToken:
        return self.value[0]


# Attention!
OperandToken.add_token_set([ControlConstructionToken])


class EntryPointToken(CompositeBaseToken):
    _TOKEN_SETS = [[EqOperatorToken, ExpressionToken]]

    @property
    def expression(self) -> ExpressionToken:
        return self.value[1]
