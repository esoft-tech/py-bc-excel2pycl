from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ExpressionToken, AmpersandToken, DateControlConstructionToken, \
    TodayControlConstructionToken, EqOperatorToken, NotEqOperatorToken, GtOperatorToken, GtOrEqualOperatorToken, \
    LtOperatorToken, LtOrEqualOperatorToken, PercentToken, OneLeftOperandExpressionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ExpressionTokenTranslator(AbstractTranslator):
    _DATE_TOKENS = [DateControlConstructionToken, TodayControlConstructionToken]

    @classmethod
    def translate(cls, token: ExpressionToken | OneLeftOperandExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.operand_token_translator import OperandTokenTranslator
        from excel2pycl.src.translators.operator_sub_token_translator import OperatorSubTokenTranslator

        operator, left_operand, left_brackets, right_brackets, right_operand = token.operator, token.left_operand, \
            None, None, None

        if isinstance(token, ExpressionToken):
            left_brackets, right_brackets, right_operand = token.left_brackets, token.right_brackets, \
                  token.right_operand

        if left_operand:
            token_translator = ExpressionTokenTranslator if \
                left_operand.__class__ in [ExpressionToken, OneLeftOperandExpressionToken] \
                else OperandTokenTranslator

            left_operand = token_translator.translate(left_operand, excel, context)
            left_operand = f'({left_operand})' if left_brackets else left_operand

        if right_operand:
            token_translator = ExpressionTokenTranslator \
                if right_operand.__class__ is ExpressionToken else OperandTokenTranslator

            right_operand = token_translator.translate(right_operand, excel, context)
            right_operand = f'({right_operand})' if right_brackets else right_operand

        if operator:
            if operator.__class__ is AmpersandToken:
                left_operand = f'str({left_operand})'
                right_operand = f'str({right_operand})'

            # попытка заставить сравнение работать так, как надо
            compare_tokens = (EqOperatorToken, NotEqOperatorToken, GtOperatorToken, GtOrEqualOperatorToken,
                              LtOperatorToken, LtOrEqualOperatorToken)

            if isinstance(operator, compare_tokens) and left_operand and right_operand:
                operator = OperatorSubTokenTranslator.translate(operator, excel, context)
                return f'self._compare("{operator}", {left_operand}, {right_operand})'

            if operator.__class__ is PercentToken:
                left_operand = f'self._normalize_float_number({left_operand} / 100)'
                operator = None
            else:
                operator = OperatorSubTokenTranslator.translate(operator, excel, context)

            if isinstance(token.left_operand, OneLeftOperandExpressionToken) and \
                    isinstance(token.left_operand.operator, PercentToken):
                return f"self._normalize_float_number({left_operand or ''}{operator or ''}{right_operand or ''})"

        return f"{left_operand or ''}{operator or ''}{right_operand or ''}"
