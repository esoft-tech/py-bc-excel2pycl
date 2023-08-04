from typing import Union

from excel2pycl.src.cell import Cell
from excel2pycl.src.tokens.composite_base_token import CompositeBaseToken
from excel2pycl.src.tokens.recursive_composite_base_token import RecursiveCompositeBaseToken, CLS
from excel2pycl.src.tokens.regexp_tokens import DayKeywordToken, LeftKeywordToken, MaxKeywordToken, MidKeywordToken, \
    MinKeywordToken, \
    ErrorKeywordToken, RightKeywordToken, SumKeywordToken, IfKeywordToken, CellIdentifierToken, \
    CellIdentifierRangeToken, \
    MatrixOfCellIdentifiersToken, EqOperatorToken, NotEqOperatorToken, GtOperatorToken, GtOrEqualOperatorToken, \
    LtOperatorToken, LtOrEqualOperatorToken, PlusOperatorToken, MinusOperatorToken, MultiplicationOperatorToken, \
    DivOperatorToken, LiteralToken, BracketStartToken, BracketFinishToken, SeparatorToken, VlookupKeywordToken, \
    AverageKeywordToken, RoundKeywordToken, OrKeywordToken, AndKeywordToken, AmpersandToken, YearKeywordToken, \
    DateKeywordToken, DifKeywordToken, EoKeywordToken, MonthKeywordToken, EKeywordToken, \
    XKeywordToken, MatchKeywordToken, SKeywordToken, AddressKeywordToken, NetworkDaysKeywordToken


class SumIfKeywordToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumKeywordToken, IfKeywordToken]]


class IfErrorKeywordToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfKeywordToken, ErrorKeywordToken]]


class DateDifKeywordToken(CompositeBaseToken):
    _TOKEN_SETS = [[DateKeywordToken, DifKeywordToken]]


class EoMonthKeywordToken(CompositeBaseToken):
    _TOKEN_SETS = [[EoKeywordToken, MonthKeywordToken]]


class EDateKeywordToken(CompositeBaseToken):
    _TOKEN_SETS = [[EKeywordToken, DateKeywordToken]]


class XMatchKeywordToken(CompositeBaseToken):
    _TOKEN_SETS = [[XKeywordToken, MatchKeywordToken]]


class SimilarCellToken(CompositeBaseToken):
    _TOKEN_SETS = [[CellIdentifierToken], [CellIdentifierRangeToken], [MatrixOfCellIdentifiersToken]]

    @property
    def cell(self) -> Cell:
        return self.value[0].cell if self.value[0].__class__ == CellIdentifierToken else None

    @property
    def range(self) -> CellIdentifierRangeToken:
        return self.value[0] if self.value[0].__class__ == CellIdentifierRangeToken else None

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken:
        return self.value[0] if self.value[0].__class__ == MatrixOfCellIdentifiersToken else None


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


class AmpersandOperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[AmpersandToken]]

    @property
    def operator(self):
        return self.value[0]


class OperandToken(CompositeBaseToken):
    _TOKEN_SETS = [[LiteralToken], [CellIdentifierToken], [CellIdentifierRangeToken], [MatrixOfCellIdentifiersToken]]

    @property
    def cell(self) -> Cell:
        return self.value[0].cell if self.value[0].__class__ == CellIdentifierToken else None

    @property
    def range(self) -> CellIdentifierRangeToken:
        return self.value[0] if self.value[0].__class__ == CellIdentifierRangeToken else None

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken:
        return self.value[0] if self.value[0].__class__ == MatrixOfCellIdentifiersToken else None

    @property
    def literal(self) -> int or float or str or bool:
        return self.value[0].value if self.value[0].__class__ == LiteralToken else None

    @property
    def control_construction(self):
        return self.value[0] if self.value[0].__class__ == ControlConstructionToken else None


class OperatorToken(CompositeBaseToken):
    _TOKEN_SETS = [[ArithmeticOperatorToken], [LogicalOperatorToken], [AmpersandOperatorToken]]

    @property
    def operator(self):
        return self.value[0].operator


class ExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [[OperandToken, OperatorToken, CLS], [OneOperandArithmeticOperatorToken, CLS],
                   [BracketStartToken, CLS, BracketFinishToken, OperatorToken, CLS],
                   [BracketStartToken, CLS, BracketFinishToken], [OperandToken]]

    @property
    def left_brackets(self) -> bool:
        return self.value[0].__class__ == BracketStartToken and len(self.value) == 5

    @property
    def right_brackets(self) -> bool:
        return self.value[0].__class__ in [OneOperandArithmeticOperatorToken, BracketStartToken] and len(
            self.value) in [2, 3]

    @property
    def operator(self):
        return self.value[0].operator if self.value[0].__class__ is OperatorToken or self.value[
            0].__class__ is OneOperandArithmeticOperatorToken else self.value[1].operator if len(
            self.value) == 3 and self.value[1].__class__ is OperatorToken else self.value[3].operator if len(
            self.value) == 5 and self.value[3].__class__ is OperatorToken else None

    @property
    def left_operand(self):
        return self.value[0] if len(self.value) in [1, 3] and self.value[0].__class__ is OperandToken else self.value[
            1] if len(self.value) == 5 and self.value[1].__class__ is self.__class__ else None

    @property
    def right_operand(self):
        return self.value[2] if len(self.value) == 3 and self.value[2].__class__ is self.__class__ else self.value[
            1] if len(self.value) in [2, 3] and self.value[1].__class__ is self.__class__ else self.value[4] if len(
            self.value) == 5 and self.value[4].__class__ is self.__class__ else None


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
    _TOKEN_SETS = [[LiteralToken, AmpersandToken, ExpressionToken], [LiteralToken], [ExpressionToken]]

    @property
    def literal(self) -> int or float or str or bool:
        return self.value[0].value if self.value[0].__class__ == LiteralToken else None

    @property
    def expression(self) -> ExpressionToken:
        return self.value[0] if self.value[0].__class__ == ExpressionToken else self.value[2] if len(
            self.value) == 3 and self.value[2].__class__ == ExpressionToken else None


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


class IfErrorControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfErrorKeywordToken,
                    BracketStartToken,
                    ExpressionToken,
                    SeparatorToken,
                    ExpressionToken,
                    BracketFinishToken]]

    @property
    def condition(self) -> ExpressionToken:
        return self.value[2]

    @property
    def when_error(self) -> ExpressionToken:
        return self.value[4]


class SumControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class SumIfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[SumIfKeywordToken, BracketStartToken, SimilarCellToken, SeparatorToken, LambdaToken,
                    BracketFinishToken],
                   [SumIfKeywordToken, BracketStartToken, SimilarCellToken, SeparatorToken, LambdaToken,
                    SeparatorToken, SimilarCellToken, BracketFinishToken]]

    @property
    def cell(self) -> Cell:
        return self.value[2].cell

    @property
    def range(self) -> CellIdentifierRangeToken:
        return self.value[2].range

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken:
        return self.value[2].matrix

    @property
    def lambda_(self) -> LambdaToken:
        return self.value[4]

    @property
    def first_cell_of_needed(self) -> Cell:
        return (self.value[6].cell if self.value[6].cell else self.value[6].range.range[0] if self.value[6].range else
                self.value[6].matrix.matrix[0] if self.value[6].matrix else None) if len(self.value) == 8 else None


class VlookupControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [
        [VlookupKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         SeparatorToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
        [VlookupKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         SeparatorToken, ExpressionToken, BracketFinishToken]]

    @property
    def lookup_value(self) -> ExpressionToken:
        return self.value[2]

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken:
        return self.value[4]

    @property
    def column_number(self) -> ExpressionToken:
        return self.value[6]

    @property
    def range_lookup(self) -> ExpressionToken:
        return self.value[8] if len(self.value) == 10 else None


class AverageControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[AverageKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class RoundControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[RoundKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken,
                    ExpressionToken, BracketFinishToken]]

    @property
    def number(self) -> ExpressionToken:
        return self.value[2]

    @property
    def num_digits(self) -> ExpressionToken:
        return self.value[4]


class DateControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[DateKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken,
                    ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken]]

    @property
    def year(self):
        return self.value[2]

    @property
    def month(self):
        return self.value[4]

    @property
    def day(self):
        return self.value[6]


class DateDifControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[DateDifKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken,
                    ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken]]

    @property
    def date_start(self):
        return self.value[2]

    @property
    def date_end(self):
        return self.value[4]

    @property
    def mode(self):
        return self.value[6]


class EoMonthControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[EoMonthKeywordToken, BracketStartToken, ExpressionToken,
                    SeparatorToken, ExpressionToken, BracketFinishToken]]

    @property
    def start_date(self) -> ExpressionToken:
        return self.value[2]

    @property
    def months(self) -> ExpressionToken:
        return self.value[4]


class EDateControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[EDateKeywordToken, BracketStartToken, ExpressionToken,
                    SeparatorToken, ExpressionToken, BracketFinishToken]]

    @property
    def start_date(self) -> ExpressionToken:
        return self.value[2]

    @property
    def months(self) -> ExpressionToken:
        return self.value[4]


class LeftControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[LeftKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
                   [LeftKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def text(self) -> ExpressionToken:
        return self.value[2]

    @property
    def num_chars(self) -> ExpressionToken:
        return self.value[4] if len(self.value) == 6 else None


class MidControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[MidKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken]]

    @property
    def text(self) -> ExpressionToken:
        return self.value[2]

    @property
    def start_num(self) -> ExpressionToken:
        return self.value[4]

    @property
    def num_chars(self) -> ExpressionToken:
        return self.value[6]


class RightControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[RightKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
                   [RightKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def text(self) -> ExpressionToken:
        return self.value[2]

    @property
    def num_chars(self) -> ExpressionToken:
        return self.value[4] if len(self.value) == 6 else None


class OrControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[OrKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class AndControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[AndKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class DayControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[DayKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def date(self):
        return self.value[2]


class MonthControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[MonthKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def date(self):
        return self.value[2]


class YearControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[YearKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def date(self):
        return self.value[2]


class MinControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[MinKeywordToken, BracketStartToken,
                    IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class MaxControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[MaxKeywordToken, BracketStartToken,
                    IterableExpressionToken, BracketFinishToken]]

    @property
    def expressions(self):
        return self.value[2].expressions


class MatchControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [
        [MatchKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         SeparatorToken, ExpressionToken, BracketFinishToken],
        [MatchKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         BracketFinishToken]]

    @property
    def lookup_value(self) -> ExpressionToken:
        return self.value[2]

    @property
    def lookup_array(self) -> MatrixOfCellIdentifiersToken:
        return self.value[4]

    @property
    def match_type(self) -> ExpressionToken:
        return self.value[6]


class XMatchControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [
        [XMatchKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         SeparatorToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
        [XMatchKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         SeparatorToken, ExpressionToken, BracketFinishToken],
        [XMatchKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, MatrixOfCellIdentifiersToken,
         BracketFinishToken]]

    @property
    def lookup_value(self) -> ExpressionToken:
        return self.value[2]

    @property
    def lookup_array(self) -> MatrixOfCellIdentifiersToken:
        return self.value[4]

    @property
    def match_mode(self) -> ExpressionToken:
        return self.value[6] if len(self.value) >= 8 else None

    @property
    def search_mode(self) -> ExpressionToken:
        return self.value[8] if len(self.value) == 10 else None

class RangeOfCellIdentifierWithConditionToken(CompositeBaseToken):
    _TOKEN_SETS = [
        [MatrixOfCellIdentifiersToken, SeparatorToken, LambdaToken],
        [MatrixOfCellIdentifiersToken, SeparatorToken, ExpressionToken]]

    @property
    def range(self) -> MatrixOfCellIdentifiersToken:
        return self.value[0]

    @property
    def condition_lambda(self) -> LambdaToken | None:
        return self.value[2] if isinstance(self.value[2], LambdaToken) else None

    @property
    def condition_expression(self) -> ExpressionToken | None:
        return None if self.condition_lambda is not None else self.value[2]


class NetworkDaysControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [
        [
            NetworkDaysKeywordToken, BracketStartToken, ExpressionToken,
            SeparatorToken, ExpressionToken, BracketFinishToken
        ],
        [
            NetworkDaysKeywordToken, BracketStartToken, ExpressionToken,
            SeparatorToken, ExpressionToken,
            SeparatorToken, MatrixOfCellIdentifiersToken,
            BracketFinishToken
        ]
    ]

    @property
    def date_start(self) -> ExpressionToken:
        return self.value[2]

    @property
    def date_end(self) -> ExpressionToken:
        return self.value[4]

    @property
    def holidays(self) -> MatrixOfCellIdentifiersToken:
        return self.value[6] if len(self.value) == 8 else None


class IterableRangeOfCellIdentifierWithConditionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS = [
        [RangeOfCellIdentifierWithConditionToken, SeparatorToken, CLS],
        [RangeOfCellIdentifierWithConditionToken]
    ]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, *kwargs)
        self._range_of_cell_identifiers_with_conditions = []

    @property
    def range_of_cell_identifiers_with_conditions(self) -> list[RangeOfCellIdentifierWithConditionToken]:
        if not self._range_of_cell_identifiers_with_conditions:
            self._range_of_cell_identifiers_with_conditions = [self.value[0]] \
                + self.value[2].range_of_cell_identifiers_with_conditions if len(self.value) == 3 else [self.value[0]]
        return self._range_of_cell_identifiers_with_conditions


class AverageIfsControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[AverageKeywordToken, IfKeywordToken, SKeywordToken, BracketStartToken, MatrixOfCellIdentifiersToken, SeparatorToken, IterableRangeOfCellIdentifierWithConditionToken, BracketFinishToken]]

    @property
    def average_range(self) -> MatrixOfCellIdentifiersToken:
        return self.value[4]

    @property
    def conditions(self) -> IterableRangeOfCellIdentifierWithConditionToken:
        return self.value[6]


class AddressControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [
        [AddressKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
        [AddressKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, SeparatorToken,
         IterableExpressionToken, BracketFinishToken]
    ]

    @property
    def row(self) -> ExpressionToken:
        return self.value[2]

    @property
    def col(self) -> ExpressionToken:
        return self.value[4]

    @property
    def expressions(self):
        return self.value[6].expressions if len(self.value) > 6 else []


class ControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS = [[IfControlConstructionToken], [SumControlConstructionToken], [SumIfControlConstructionToken],
                   [VlookupControlConstructionToken], [AverageControlConstructionToken],
                   [RoundControlConstructionToken], [OrControlConstructionToken], [AndControlConstructionToken],
                   [EDateControlConstructionToken], [EoMonthControlConstructionToken],
                   [DateDifControlConstructionToken], [YearControlConstructionToken], [MonthControlConstructionToken],
                   [DayControlConstructionToken], [MinControlConstructionToken], [MaxControlConstructionToken],
                   [IfErrorControlConstructionToken], [DateControlConstructionToken], [MatchControlConstructionToken],
                   [XMatchControlConstructionToken], [LeftControlConstructionToken], [MidControlConstructionToken],
                   [RightControlConstructionToken], [AverageIfsControlConstructionToken],
                   [AddressControlConstructionToken], [NetworkDaysControlConstructionToken]]

    @property
    def control_construction(self) -> Union[IfControlConstructionToken, SumControlConstructionToken,
                                            SumIfControlConstructionToken, VlookupControlConstructionToken,
                                            AverageControlConstructionToken, RoundControlConstructionToken,
                                            OrControlConstructionToken, AndControlConstructionToken,
                                            YearControlConstructionToken, MonthControlConstructionToken,
                                            DayControlConstructionToken,
                                            MinControlConstructionToken, MaxControlConstructionToken,
                                            IfErrorControlConstructionToken, DateControlConstructionToken,
                                            DateDifControlConstructionToken, EoMonthControlConstructionToken,
                                            EDateControlConstructionToken, MatchControlConstructionToken,
                                            XMatchControlConstructionToken, LeftControlConstructionToken,
                                            MidControlConstructionToken, RightControlConstructionToken,
                                            AverageIfsControlConstructionToken, AddressControlConstructionToken,
                                            NetworkDaysControlConstructionToken]:
        return self.value[0]


# Attention!
OperandToken.add_token_set([ControlConstructionToken])


class EntryPointToken(CompositeBaseToken):
    _TOKEN_SETS = [[EqOperatorToken, ExpressionToken]]

    @property
    def expression(self) -> ExpressionToken:
        return self.value[1]
