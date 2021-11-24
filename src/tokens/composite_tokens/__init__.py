from src.tokens.composite_base_token import CompositeBaseToken
from src.tokens.recursive_composite_base_token import RecursiveCompositeBaseToken, CLS
from src.tokens.regexp_tokens import *


class LogicalOperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[EqOperatorToken], [NotEqOperatorToken], [GtOperatorToken], [GtOrEqualOperatorToken],
                   [LtOperatorToken], [LtOrEqualOperatorToken]]


class OneOperandArithmeticOperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[PlusOperatorToken], [MinusOperatorToken]]


class ArithmeticOperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[OneOperandArithmeticOperatorToken], [MultiplicationOperatorToken], [DivOperatorToken]]


class OperandToken(CompositeBaseToken):
    _TOKEN_SETS = [[LiteralToken], [CellIdentifierToken], [CellIdentifierRangeToken], [MatrixOfCellIdentifiersToken]]


class OperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[ArithmeticOperatorToken], [LogicalOperatorToken]]


class ExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [[OperandToken, OperatorToken, CLS], [OneOperandArithmeticOperatorToken, OperandToken],
                   [BracketStartToken, CLS, BracketFinishToken], [OperandToken]]


class IterableExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [[ExpressionToken, SeparatorToken, CLS], [ExpressionToken]]


class LambdaToken(CompositeBaseToken):
    _TOKEN_SETS = [[LiteralToken, AndLambdaToken, ExpressionToken], [ExpressionToken]]


class IfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, SeparatorToken,
                    ExpressionToken, BracketFinishToken]]


class SumControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]


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


class ControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfControlConstructionToken], [SumControlConstructionToken], [SumIfControlConstructionToken],
                   [VlookupControlConstructionToken], [AverageControlConstructionToken]]


OperandToken.add_token_set([ControlConstructionToken])


class EntryPointToken(CompositeBaseToken):
    _TOKEN_SETS = [[EqOperatorToken, ExpressionToken]]
