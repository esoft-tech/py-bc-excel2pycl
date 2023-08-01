from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ExpressionToken, AmpersandToken, DateControlConstructionToken, \
    TodayControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ExpressionTokenTranslator(AbstractTranslator):
    _DATE_TOKENS = [DateControlConstructionToken, TodayControlConstructionToken]

    @classmethod
    def translate(cls, token: ExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.operand_token_translator import OperandTokenTranslator
        from excel2pycl.src.translators.operator_sub_token_translator import OperatorSubTokenTranslator
        from excel2pycl.src.utilities.helper import get_operand_source_token

        left_brackets, right_brackets, operator, left_operand, right_operand = token.left_brackets, token.right_brackets, token.operator, token.left_operand, token.right_operand
        left_operand_class = get_operand_source_token(left_operand)
        right_operand_class = get_operand_source_token(right_operand)

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

            operator = OperatorSubTokenTranslator.translate(operator, excel, context)

            if left_operand_class in cls._DATE_TOKENS and isinstance(right_operand, str) and right_operand.isdigit():
                return f'{left_operand}{operator}datetime.timedelta(days={right_operand})'
            elif right_operand_class in cls._DATE_TOKENS and isinstance(left_operand, str) and left_operand.isdigit():
                return f'datetime.timedelta(days={left_operand}){operator}{right_operand}'

        return f"{left_operand or ''}{operator or ''}{right_operand or ''}"
