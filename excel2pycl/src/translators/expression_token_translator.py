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
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ExpressionTokenTranslator(AbstractTranslator[ExpressionToken]):
    _DATE_TOKENS: ClassVar[list] = [DateControlConstructionToken, TodayControlConstructionToken]

    @classmethod
    def translate(cls, token: ExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.operand_token_translator import OperandTokenTranslator
        from excel2pycl.src.translators.operator_sub_token_translator import OperatorSubTokenTranslator

        left_brackets, right_brackets, operator, left_operand, right_operand = (
            token.left_brackets,
            token.right_brackets,
            token.operator,
            token.left_operand,
            token.right_operand,
        )
        translated_left_operand = ""
        translated_right_operand = ""
        translated_operator = ""

        if left_operand:
            token_translator = (
                ExpressionTokenTranslator if left_operand.__class__ is ExpressionToken else OperandTokenTranslator
            )
            translated_right_operand = token_translator.translate(left_operand, excel, context)  # type: ignore
            translated_left_operand = f"({translated_right_operand})" if left_brackets else translated_right_operand

        if right_operand:
            token_translator = (
                ExpressionTokenTranslator if right_operand.__class__ is ExpressionToken else OperandTokenTranslator
            )
            translated_right_operand = token_translator.translate(right_operand, excel, context)  # type: ignore
            translated_right_operand = f"({translated_right_operand})" if right_brackets else translated_right_operand

        if operator:
            if operator.__class__ is AmpersandToken:
                translated_left_operand = f"str({left_operand})"
                translated_right_operand = f"str({right_operand})"

            # попытка заставить сравнение работать так, как надо
            compare_tokens = (
                EqOperatorToken,
                NotEqOperatorToken,
                GtOperatorToken,
                GtOrEqualOperatorToken,
                LtOperatorToken,
                LtOrEqualOperatorToken,
            )

            if isinstance(operator, compare_tokens) and translated_left_operand and translated_left_operand:
                translated_operator = OperatorSubTokenTranslator.translate(operator, excel, context)
                return f'self._compare("{translated_operator}", {translated_left_operand}, {translated_right_operand})'

            translated_operator = OperatorSubTokenTranslator.translate(operator, excel, context)

        return f"{translated_left_operand}{translated_operator}{translated_right_operand}"
