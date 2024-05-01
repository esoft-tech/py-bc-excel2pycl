from typing import ClassVar

from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import (
    AmpersandToken,
    DateControlConstructionToken,
    EqOperatorToken,
    ExpressionToken,
    GtOperatorToken,
    GtOrEqualOperatorToken,
    LtOperatorToken,
    LtOrEqualOperatorToken,
    NotEqOperatorToken,
    TodayControlConstructionToken,
)
from excel2pycl.src.tokens.composite_tokens import OperandToken
from excel2pycl.src.tokens.regexp_tokens import (
    DivOperatorToken,
    MinusOperatorToken,
    MultiplicationOperatorToken,
    PlusOperatorToken,
)
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.translators.operand_token_translator import OperandTokenTranslator
from excel2pycl.src.translators.operator_sub_token_translator import OperatorSubTokenTranslator


class ExpressionTokenTranslator(AbstractTranslator[ExpressionToken]):
    _DATE_TOKENS: ClassVar[list] = [DateControlConstructionToken, TodayControlConstructionToken]

    @classmethod
    def _translate_operand(
        cls,
        operand: ExpressionToken | OperandToken | None,
        brackets: bool,  # noqa: FBT001
        excel: Excel,
        context: Context,
    ) -> str:
        translated_operand = ""

        if operand:
            if isinstance(operand, ExpressionToken):
                translated_operand = ExpressionTokenTranslator.translate(operand, excel, context)
            else:
                translated_operand = OperandTokenTranslator.translate(operand, excel, context)

            return f"({translated_operand})" if brackets else translated_operand

        return translated_operand

    @classmethod
    def _translate_operator(
        cls,
        operator: MultiplicationOperatorToken
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
        | None,
        translated_left_operand: str,
        translated_right_operand: str,
        excel: Excel,
        context: Context,
    ) -> str:
        translated_operator = ""

        if operator:
            if isinstance(operator, AmpersandToken):
                translated_left_operand = f"str({translated_left_operand})"
                translated_right_operand = f"str({translated_right_operand})"

            # попытка заставить сравнение работать так, как надо
            compare_tokens = (
                EqOperatorToken,
                NotEqOperatorToken,
                GtOperatorToken,
                GtOrEqualOperatorToken,
                LtOperatorToken,
                LtOrEqualOperatorToken,
            )

            if isinstance(operator, compare_tokens) and translated_left_operand and translated_right_operand:
                translated_operator = OperatorSubTokenTranslator.translate(operator, excel, context)
                return f'self._compare("{translated_operator}", {translated_left_operand}, {translated_right_operand})'

            translated_operator = OperatorSubTokenTranslator.translate(operator, excel, context)

        return translated_operator

    @classmethod
    def translate(cls, token: ExpressionToken, excel: Excel, context: Context) -> str:
        left_brackets, right_brackets, operator, left_operand, right_operand = (
            token.left_brackets,
            token.right_brackets,
            token.operator,
            token.left_operand,
            token.right_operand,
        )

        translated_left_operand = cls._translate_operand(left_operand, left_brackets, excel, context)

        translated_right_operand = cls._translate_operand(right_operand, right_brackets, excel, context)

        translated_operator = cls._translate_operator(
            operator,
            translated_left_operand,
            translated_right_operand,
            excel,
            context,
        )

        if "self._compare" in translated_operator:
            return translated_operator

        return f"{translated_left_operand}{translated_operator}{translated_right_operand}"
