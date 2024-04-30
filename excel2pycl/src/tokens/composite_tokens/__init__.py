from collections.abc import Iterable
from typing import Any, ClassVar, TypeAlias, TypeVar, Union, cast

from excel2pycl.src.cell import Cell
from excel2pycl.src.tokens.composite_base_token import CompositeBaseToken
from excel2pycl.src.tokens.recursive_composite_base_token import CLS, RecursiveCompositeBaseToken
from excel2pycl.src.tokens.regexp_tokens import (
    AddressKeywordToken,
    AmpersandToken,
    AndKeywordToken,
    AverageIfSKeywordToken,
    AverageKeywordToken,
    BracketFinishToken,
    BracketStartToken,
    CellIdentifierRangeToken,
    CellIdentifierToken,
    ColumnKeywordToken,
    CountBlankKeywordToken,
    CountIfSKeywordToken,
    CountKeywordToken,
    DateDifKeywordToken,
    DateKeywordToken,
    DayKeywordToken,
    DivOperatorToken,
    EDateKeywordToken,
    EoMonthKeywordToken,
    EqOperatorToken,
    GtOperatorToken,
    GtOrEqualOperatorToken,
    IfErrorKeywordToken,
    IfKeywordToken,
    IfsKeywordToken,
    IndexKeywordToken,
    LeftKeywordToken,
    LiteralToken,
    LtOperatorToken,
    LtOrEqualOperatorToken,
    MatchKeywordToken,
    MatrixOfCellIdentifiersToken,
    MaxKeywordToken,
    MidKeywordToken,
    MinKeywordToken,
    MinusOperatorToken,
    MonthKeywordToken,
    MultiplicationOperatorToken,
    NetworkDaysKeywordToken,
    NotEqOperatorToken,
    OrKeywordToken,
    PatternToken,
    PlusOperatorToken,
    RightKeywordToken,
    RoundKeywordToken,
    SearchKeywordToken,
    SeparatorToken,
    SumIfKeywordToken,
    SumIfSKeywordToken,
    SumKeywordToken,
    TodayKeywordToken,
    VlookupKeywordToken,
    XMatchKeywordToken,
    YearKeywordToken,
)


class SimilarCellToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [CellIdentifierToken],
        [CellIdentifierRangeToken],
        [MatrixOfCellIdentifiersToken],
    ]

    @property
    def cell(self) -> Cell | None:
        return self.value[0].cell if self.value[0].__class__ == CellIdentifierToken else None

    @property
    def range(self) -> CellIdentifierRangeToken | None:
        return self.value[0] if self.value[0].__class__ == CellIdentifierRangeToken else None

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken | None:
        return self.value[0] if self.value[0].__class__ == MatrixOfCellIdentifiersToken else None


class LogicalOperatorToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [EqOperatorToken],
        [NotEqOperatorToken],
        [GtOperatorToken],
        [GtOrEqualOperatorToken],
        [LtOperatorToken],
        [LtOrEqualOperatorToken],
    ]

    @property
    def operator(
        self,
    ) -> (
        EqOperatorToken
        | NotEqOperatorToken
        | GtOperatorToken
        | GtOperatorToken
        | GtOrEqualOperatorToken
        | LtOperatorToken
        | LtOrEqualOperatorToken
    ):
        return cast(
            (
                EqOperatorToken
                | NotEqOperatorToken
                | GtOperatorToken
                | GtOperatorToken
                | GtOrEqualOperatorToken
                | LtOperatorToken
                | LtOrEqualOperatorToken
            ),
            self.value[0],
        )


class OneOperandArithmeticOperatorToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[PlusOperatorToken], [MinusOperatorToken]]

    @property
    def operator(self) -> PlusOperatorToken | MinusOperatorToken:
        return cast(PlusOperatorToken | MinusOperatorToken, self.value[0])


class ArithmeticOperatorToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [OneOperandArithmeticOperatorToken],
        [MultiplicationOperatorToken],
        [DivOperatorToken],
    ]

    @property
    def operator(
        self,
    ) -> MultiplicationOperatorToken | DivOperatorToken | PlusOperatorToken | MinusOperatorToken:
        return self.value[0].operator if self.value[0].__class__ is OneOperandArithmeticOperatorToken else self.value[0]  # type: ignore [no-any-return]


class AmpersandOperatorToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[AmpersandToken]]

    @property
    def operator(self) -> AmpersandToken:
        return cast(AmpersandToken, self.value[0])


class OperandToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [PatternToken],
        [LiteralToken],
        [CellIdentifierToken],
        [CellIdentifierRangeToken],
        [MatrixOfCellIdentifiersToken],
    ]

    @property
    def pattern(self) -> PatternToken | None:
        return self.value[0] if self.value[0].__class__ == PatternToken else None

    @property
    def cell(self) -> Cell | None:
        return self.value[0].cell if self.value[0].__class__ == CellIdentifierToken else None

    @property
    def range(self) -> CellIdentifierRangeToken | None:
        return self.value[0] if self.value[0].__class__ == CellIdentifierRangeToken else None

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken | None:
        return self.value[0] if self.value[0].__class__ == MatrixOfCellIdentifiersToken else None

    @property
    def literal(self) -> int | float | str | bool | None:
        return self.value[0].value if self.value[0].__class__ == LiteralToken else None

    @property
    def control_construction(self) -> "ControlConstructionCompositeBaseToken | None":
        # ToDo: Разобраться, что здесь возвращается в self.value[0]
        return self.value[0] if self.value[0].__class__ == ControlConstructionCompositeBaseToken else None


class OperatorToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[ArithmeticOperatorToken], [LogicalOperatorToken], [AmpersandOperatorToken]]

    @property
    def operator(self) -> ArithmeticOperatorToken | LogicalOperatorToken | AmpersandOperatorToken:
        return cast(ArithmeticOperatorToken | LogicalOperatorToken | AmpersandOperatorToken, self.value[0].operator)


class ExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [OperandToken, OperatorToken, CLS],
        [OneOperandArithmeticOperatorToken, CLS],
        [BracketStartToken, CLS, BracketFinishToken, OperatorToken, CLS],
        [BracketStartToken, CLS, BracketFinishToken],
        [OperandToken],
    ]

    @property
    def left_brackets(self) -> bool:
        return self.value[0].__class__ == BracketStartToken and len(self.value) == 5

    @property
    def right_brackets(self) -> bool:
        return self.value[0].__class__ in [OneOperandArithmeticOperatorToken, BracketStartToken] and len(
            self.value,
        ) in [2, 3]

    @property
    def operator(
        self,
    ) -> (
        MultiplicationOperatorToken
        | DivOperatorToken
        | EqOperatorToken
        | NotEqOperatorToken
        | GtOperatorToken
        | GtOperatorToken
        | GtOrEqualOperatorToken
        | LtOperatorToken
        | LtOrEqualOperatorToken
        | AmpersandToken
        | PlusOperatorToken
        | MinusOperatorToken
        | None
    ):
        return (
            self.value[0].operator
            if self.value[0].__class__ is OperatorToken or self.value[0].__class__ is OneOperandArithmeticOperatorToken
            else self.value[1].operator
            if len(self.value) == 3 and self.value[1].__class__ is OperatorToken
            else self.value[3].operator
            if len(self.value) == 5 and self.value[3].__class__ is OperatorToken
            else None
        )

    @property
    def left_operand(self) -> "ExpressionToken | OperandToken | None":
        return (
            self.value[0]
            if len(self.value) in [1, 3] and self.value[0].__class__ is OperandToken
            else self.value[1]
            if len(self.value) == 5 and self.value[1].__class__ is self.__class__
            else None
        )

    @property
    def right_operand(self) -> "ExpressionToken | None":
        return (
            self.value[2]
            if len(self.value) == 3 and self.value[2].__class__ is self.__class__
            else self.value[1]
            if len(self.value) in [2, 3] and self.value[1].__class__ is self.__class__
            else self.value[4]
            if len(self.value) == 5 and self.value[4].__class__ is self.__class__
            else None
        )


class IterableExpressionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[ExpressionToken, SeparatorToken, CLS], [ExpressionToken]]

    def __init__(self, *args: tuple, **kwargs: dict) -> None:
        # ToDo: Проверить, что __init__ получает все аргументы
        super().__init__(*args, **kwargs)  # type: ignore
        self._expressions: list = []

    @property
    def expressions(self) -> list:
        if not self._expressions:
            self._expressions = [self.value[0], *self.value[2].expressions] if len(self.value) == 3 else [self.value[0]]
        return self._expressions


class LambdaToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [LiteralToken, AmpersandToken, ExpressionToken],
        [LiteralToken],
        [ExpressionToken],
    ]

    @property
    def literal(self) -> int | float | str | bool | None:
        return self.value[0].value if self.value[0].__class__ == LiteralToken else None

    @property
    def expression(self) -> ExpressionToken | None:
        return (
            self.value[0]
            if self.value[0].__class__ == ExpressionToken
            else self.value[2]
            if len(self.value) == 3 and self.value[2].__class__ == ExpressionToken
            else None
        )


class IfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            IfKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [IfKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
    ]

    @property
    def condition(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def when_true(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def when_false(self) -> ExpressionToken | None:
        return self.value[6] if len(self.value) == 8 else None


class IfErrorControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [IfErrorKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
    ]

    @property
    def condition(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def when_error(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])


class SumControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [SumKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class SumIfControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [SumIfKeywordToken, BracketStartToken, SimilarCellToken, SeparatorToken, LambdaToken, BracketFinishToken],
        [
            SumIfKeywordToken,
            BracketStartToken,
            SimilarCellToken,
            SeparatorToken,
            LambdaToken,
            SeparatorToken,
            SimilarCellToken,
            BracketFinishToken,
        ],
    ]

    @property
    def cell(self) -> Cell:
        return cast(Cell, self.value[2].cell)

    @property
    def range(self) -> CellIdentifierRangeToken:
        return cast(CellIdentifierRangeToken, self.value[2].range)

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[2].matrix)

    @property
    def lambda_(self) -> LambdaToken:
        return cast(LambdaToken, self.value[4])

    @property
    def first_cell_of_needed(self) -> Cell | None:
        return (
            (
                self.value[6].cell
                if self.value[6].cell
                else self.value[6].range.range[0]
                if self.value[6].range
                else self.value[6].matrix.matrix[0]
                if self.value[6].matrix
                else None
            )
            if len(self.value) == 8
            else None
        )


class VlookupControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            VlookupKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            VlookupKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def lookup_value(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[4])

    @property
    def column_number(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[6])

    @property
    def range_lookup(self) -> ExpressionToken | None:
        return self.value[8] if len(self.value) == 10 else None


class AverageControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [AverageKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class RoundControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [RoundKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
    ]

    @property
    def number(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def num_digits(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])


class DateControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            DateKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def year(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def month(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def day(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[6])


class DateDifControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            DateDifKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def date_start(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def date_end(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def mode(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[6])


class EoMonthControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [EoMonthKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
    ]

    @property
    def start_date(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def months(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])


class EDateControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [EDateKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
    ]

    @property
    def start_date(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def months(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])


class LeftControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [LeftKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
        [LeftKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken],
    ]

    @property
    def text(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def num_chars(self) -> ExpressionToken | None:
        return self.value[4] if len(self.value) == 6 else None


class SearchControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            SearchKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            SearchKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def find_text(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def within_text(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def start_num(self) -> ExpressionToken | None:
        return self.value[6] if len(self.value) == 8 else None


class MidControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            MidKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def text(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def start_num(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def num_chars(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[6])


class RightControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [RightKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
        [RightKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken],
    ]

    @property
    def text(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def num_chars(self) -> ExpressionToken | None:
        return self.value[4] if len(self.value) == 6 else None


class OrControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [OrKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class AndControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [AndKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class DayControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[DayKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def date(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])


class MonthControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[MonthKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def date(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])


class YearControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[YearKeywordToken, BracketStartToken, ExpressionToken, BracketFinishToken]]

    @property
    def date(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])


class MinControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [MinKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class MaxControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [MaxKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class MatchControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            MatchKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            MatchKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            BracketFinishToken,
        ],
    ]

    @property
    def lookup_value(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def lookup_array(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[4])

    @property
    def match_type(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[6])


class XMatchControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            XMatchKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            XMatchKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            XMatchKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            BracketFinishToken,
        ],
    ]

    @property
    def lookup_value(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def lookup_array(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[4])

    @property
    def match_mode(self) -> ExpressionToken | None:
        return self.value[6] if len(self.value) >= 8 else None

    @property
    def search_mode(self) -> ExpressionToken | None:
        return self.value[8] if len(self.value) == 10 else None


class RangeOfCellIdentifierWithConditionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [MatrixOfCellIdentifiersToken, SeparatorToken, LambdaToken],
        [MatrixOfCellIdentifiersToken, SeparatorToken, ExpressionToken],
    ]

    @property
    def range(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[0])

    @property
    def condition_lambda(self) -> LambdaToken | None:
        return self.value[2] if isinstance(self.value[2], LambdaToken) else None

    @property
    def condition_expression(self) -> ExpressionToken | None:
        return None if self.condition_lambda is not None else self.value[2]


class NetworkDaysControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            NetworkDaysKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            NetworkDaysKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            MatrixOfCellIdentifiersToken,
            BracketFinishToken,
        ],
    ]

    @property
    def date_start(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def date_end(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def holidays(self) -> MatrixOfCellIdentifiersToken | None:
        return self.value[6] if len(self.value) == 8 else None


class IterableRangeOfCellIdentifierWithConditionToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [RangeOfCellIdentifierWithConditionToken, SeparatorToken, CLS],
        [RangeOfCellIdentifierWithConditionToken],
    ]

    def __init__(self, *args: tuple, **kwargs: dict) -> None:
        # ToDo: Проверить работу __init__
        super().__init__(*args, **kwargs)  # type: ignore [arg-type]
        self._range_of_cell_identifiers_with_conditions: list = []

    @property
    def range_of_cell_identifiers_with_conditions(self) -> list[RangeOfCellIdentifierWithConditionToken]:
        if not self._range_of_cell_identifiers_with_conditions:
            self._range_of_cell_identifiers_with_conditions = (
                [self.value[0], *self.value[2].range_of_cell_identifiers_with_conditions]
                if len(self.value) == 3
                else [self.value[0]]
            )
        return self._range_of_cell_identifiers_with_conditions


class AverageIfsControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            AverageIfSKeywordToken,
            BracketStartToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            IterableRangeOfCellIdentifierWithConditionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def average_range(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[2])

    @property
    def conditions(self) -> IterableRangeOfCellIdentifierWithConditionToken:
        return cast(IterableRangeOfCellIdentifierWithConditionToken, self.value[4])


class CountBlankControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [CountBlankKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class IfsControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [IfsKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def expressions(self) -> list:
        return cast(list, self.value[2].expressions)


class CountIfsControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            CountIfSKeywordToken,
            BracketStartToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            LambdaToken,
            SeparatorToken,
            IterableRangeOfCellIdentifierWithConditionToken,
            BracketFinishToken,
        ],
        [
            CountIfSKeywordToken,
            BracketStartToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            LambdaToken,
            BracketFinishToken,
        ],
    ]

    @property
    def count_range(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[2])

    @property
    def count_condition(self) -> LambdaToken:
        return cast(LambdaToken, self.value[4])

    @property
    def conditions(self) -> IterableRangeOfCellIdentifierWithConditionToken | None:
        return self.value[6] if len(self.value) > 6 else None


class AddressControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [AddressKeywordToken, BracketStartToken, ExpressionToken, SeparatorToken, ExpressionToken, BracketFinishToken],
        [
            AddressKeywordToken,
            BracketStartToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            IterableExpressionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def row(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[2])

    @property
    def col(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def expressions(self) -> list:
        return cast(list, self.value[6].expressions) if len(self.value) > 6 else []


class CountControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [CountKeywordToken, BracketStartToken, IterableExpressionToken, BracketFinishToken],
    ]

    @property
    def matrices(self) -> list[MatrixOfCellIdentifiersToken]:
        return [
            expression.left_operand.matrix
            for expression in self.value[2].expressions
            if hasattr(expression.left_operand, "matrix") and expression.left_operand.matrix is not None
        ]

    @property
    def arg_cells(self) -> list[CellIdentifierToken]:
        return [
            expression.left_operand.value[0]
            for expression in self.value[2].expressions
            if isinstance(expression.left_operand.value[0], CellIdentifierToken)
        ]

    @property
    def expressions(self) -> list:
        return [
            expression
            for expression in self.value[2].expressions
            if isinstance(expression.left_operand.value[0], LiteralToken)
        ]


class ColumnControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [ColumnKeywordToken, BracketStartToken, BracketFinishToken],
        [ColumnKeywordToken, BracketStartToken, CellIdentifierToken, BracketFinishToken],
        [ColumnKeywordToken, BracketStartToken, MatrixOfCellIdentifiersToken, BracketFinishToken],
    ]

    @property
    def cell(self) -> CellIdentifierToken | None:
        return self.value[2] if len(self.value) > 2 and isinstance(self.value[2], CellIdentifierToken) else None

    @property
    def matrix(self) -> MatrixOfCellIdentifiersToken | None:
        return (
            self.value[2] if len(self.value) > 2 and isinstance(self.value[2], MatrixOfCellIdentifiersToken) else None
        )


class TodayControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[TodayKeywordToken, BracketStartToken, BracketFinishToken]]


class SumIfsControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            SumIfSKeywordToken,
            BracketStartToken,
            MatrixOfCellIdentifiersToken,
            SeparatorToken,
            IterableRangeOfCellIdentifierWithConditionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def sum_range(self) -> MatrixOfCellIdentifiersToken:
        return cast(MatrixOfCellIdentifiersToken, self.value[2])

    @property
    def conditions(self) -> IterableRangeOfCellIdentifierWithConditionToken:
        return cast(IterableRangeOfCellIdentifierWithConditionToken, self.value[4])


class IterableMatrixOfCellIdentifiersToken(RecursiveCompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [BracketStartToken, MatrixOfCellIdentifiersToken, SeparatorToken, CLS, BracketFinishToken],
        [MatrixOfCellIdentifiersToken, SeparatorToken, CLS],
        [MatrixOfCellIdentifiersToken],
    ]

    def __init__(self, *args: tuple, **kwargs: dict) -> None:
        super().__init__(*args, **kwargs)  # type: ignore [arg-type]
        self._matrix_list: list = []

    @property
    def matrix_list(self) -> list:
        if not self._matrix_list:
            if len(self.value) == 5:
                self._matrix_list = [self.value[1], *self.value[3].matrix_list]
            elif len(self.value) == 3:
                self._matrix_list = [self.value[0], *self.value[2].matrix_list]
            else:
                self._matrix_list = [self.value[0]]
        return self._matrix_list


class IndexControlConstructionToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [
            IndexKeywordToken,
            BracketStartToken,
            IterableMatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            IndexKeywordToken,
            BracketStartToken,
            IterableMatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
        [
            IndexKeywordToken,
            BracketStartToken,
            IterableMatrixOfCellIdentifiersToken,
            SeparatorToken,
            ExpressionToken,
            BracketFinishToken,
        ],
    ]

    @property
    def matrix_list(self) -> list:
        return cast(list, self.value[2].matrix_list)

    @property
    def row_number(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[4])

    @property
    def column_number(self) -> ExpressionToken | None:
        return self.value[6] if len(self.value) >= 8 else None

    @property
    def area_number(self) -> ExpressionToken | None:
        return self.value[8] if len(self.value) == 10 else None


class ControlConstructionCompositeBaseToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [
        [IfControlConstructionToken],
        [SumIfControlConstructionToken],
        [SumControlConstructionToken],
        [VlookupControlConstructionToken],
        [AverageControlConstructionToken],
        [RoundControlConstructionToken],
        [OrControlConstructionToken],
        [AndControlConstructionToken],
        [EDateControlConstructionToken],
        [EoMonthControlConstructionToken],
        [DateDifControlConstructionToken],
        [YearControlConstructionToken],
        [MonthControlConstructionToken],
        [DayControlConstructionToken],
        [MinControlConstructionToken],
        [MaxControlConstructionToken],
        [IfErrorControlConstructionToken],
        [DateControlConstructionToken],
        [MatchControlConstructionToken],
        [XMatchControlConstructionToken],
        [LeftControlConstructionToken],
        [MidControlConstructionToken],
        [RightControlConstructionToken],
        [AverageIfsControlConstructionToken],
        [CountBlankControlConstructionToken],
        [IfsControlConstructionToken],
        [TodayControlConstructionToken],
        [SearchControlConstructionToken],
        [AddressControlConstructionToken],
        [CountIfsControlConstructionToken],
        [CountControlConstructionToken],
        [NetworkDaysControlConstructionToken],
        [ColumnControlConstructionToken],
        [SumIfsControlConstructionToken],
        [IndexControlConstructionToken],
    ]

    @property
    def control_construction(
        self,
    ) -> Union[
        IfControlConstructionToken,
        SumIfControlConstructionToken,
        SumControlConstructionToken,
        VlookupControlConstructionToken,
        AverageControlConstructionToken,
        RoundControlConstructionToken,
        OrControlConstructionToken,
        AndControlConstructionToken,
        YearControlConstructionToken,
        MonthControlConstructionToken,
        DayControlConstructionToken,
        MinControlConstructionToken,
        MaxControlConstructionToken,
        IfErrorControlConstructionToken,
        DateControlConstructionToken,
        DateDifControlConstructionToken,
        EoMonthControlConstructionToken,
        EDateControlConstructionToken,
        MatchControlConstructionToken,
        XMatchControlConstructionToken,
        LeftControlConstructionToken,
        MidControlConstructionToken,
        RightControlConstructionToken,
        CountBlankControlConstructionToken,
        TodayControlConstructionToken,
        AverageIfsControlConstructionToken,
        SearchControlConstructionToken,
        CountIfsControlConstructionToken,
        IfsControlConstructionToken,
        AddressControlConstructionToken,
        CountControlConstructionToken,
        NetworkDaysControlConstructionToken,
        ColumnControlConstructionToken,
        SumIfsControlConstructionToken,
        IndexControlConstructionToken,
    ]:
        return self.value[0]  # type: ignore


# Attention!
OperandToken.add_token_set([ControlConstructionCompositeBaseToken])


class EntryPointToken(CompositeBaseToken):
    _TOKEN_SETS: ClassVar[list[list]] = [[EqOperatorToken, ExpressionToken]]

    @property
    def expression(self) -> ExpressionToken:
        return cast(ExpressionToken, self.value[1])
