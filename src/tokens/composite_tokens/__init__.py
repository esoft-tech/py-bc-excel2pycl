from src.tokens.composite_base_token import CompositeBaseToken
from src.tokens.recursive_composite_base_token import RecursiveCompositeBaseToken, CLS
from src.tokens.regexp_tokens import *


class OperandToken(CompositeBaseToken):
    _TOKEN_SETS = [[LiteralToken], [CellIdentifierToken], [SetOfCellIdentifiersToken], [MatrixOfCellIdentifiersToken]]


class ExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [[OperandToken, OperatorToken, CLS], [OperatorToken, OperandToken],
                   [BracketStartToken, CLS, BracketFinishToken], [OperandToken]]


class IterableExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [[ExpressionToken, SeparatorToken, CLS], [ExpressionToken]]


class LambdaToken(CompositeBaseToken):
    _TOKEN_SETS = [[LiteralToken, AndToken, ExpressionToken], [ExpressionToken]]


class IfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, SeparatorToken,
                    ExpressionToken, BracketFinishToken]]


class SumControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]


class SumIfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumIfKeywordToken, BracketStartToken, SetOfCellIdentifiersToken, SeparatorToken, LambdaToken,
                    BracketFinishToken],
                   [SumIfKeywordToken, BracketStartToken, SetOfCellIdentifiersToken, SeparatorToken, LambdaToken,
                    SeparatorToken, SetOfCellIdentifiersToken, BracketFinishToken]]


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
