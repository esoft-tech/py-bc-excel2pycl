from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ExpressionToken, AmpersandToken, DateControlConstructionToken, \
    TodayControlConstructionToken, EqOperatorToken, NotEqOperatorToken, GtOperatorToken, GtOrEqualOperatorToken, \
    LtOperatorToken, LtOrEqualOperatorToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ExpressionTokenTranslator(AbstractTranslator):
    _DATE_TOKENS = [DateControlConstructionToken, TodayControlConstructionToken]

    @classmethod
    def translate(cls, token: ExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.operand_token_translator import OperandTokenTranslator
        from excel2pycl.src.translators.operator_sub_token_translator import OperatorSubTokenTranslator

        left_brackets, right_brackets, operator, left_operand, right_operand = token.left_brackets, token.right_brackets, token.operator, token.left_operand, token.right_operand

        if left_operand:
            token_translator = ExpressionTokenTranslator if left_operand.__class__ is ExpressionToken else OperandTokenTranslator
            left_operand = token_translator.translate(left_operand, excel, context)
            left_operand = f'({left_operand})' if left_brackets else left_operand

        if right_operand:
            token_translator = ExpressionTokenTranslator if right_operand.__class__ is ExpressionToken else OperandTokenTranslator
            right_operand = token_translator.translate(right_operand, excel, context)
            right_operand = f'({right_operand})' if right_brackets else right_operand

        if operator:
            if operator.__class__ is AmpersandToken:
                left_operand = f'str({left_operand})'
                right_operand = f'str({right_operand})'

            # попытка заставить сравнение работать так, как надо
            compare_tokens = {EqOperatorToken, NotEqOperatorToken, GtOperatorToken, GtOrEqualOperatorToken,
                              LtOperatorToken, LtOrEqualOperatorToken}

            if operator.__class__ in compare_tokens and left_operand and right_operand:
                operator = OperatorSubTokenTranslator.translate(operator, excel, context)
                return f'self._especial_compare("{operator}", {left_operand}, {right_operand})'

            operator = OperatorSubTokenTranslator.translate(operator, excel, context)

        return f"{left_operand or ''}{operator or ''}{right_operand or ''}"
