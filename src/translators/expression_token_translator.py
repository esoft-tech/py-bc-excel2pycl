from src.context import Context
from src.excel import Excel
from src.tokens import ExpressionToken
from src.translators.abstract_translator import AbstractTranslator


class ExpressionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: ExpressionToken, excel: Excel, context: Context) -> str:
        from src.translators.operand_token_translator import OperandTokenTranslator
        from src.translators.operator_sub_token_translator import OperatorSubTokenTranslator

        in_brackets, operator, left_operand, right_operand = token.in_brackets, token.operator, token.left_operand, token.right_operand

        result = []
        if in_brackets:
            result.append('(')

        if left_operand:
            result.append(OperandTokenTranslator.translate(left_operand, excel, context))

        if operator:
            result.append(OperatorSubTokenTranslator.translate(operator, excel, context))

        if right_operand:
            result.append(ExpressionTokenTranslator.translate(right_operand, excel,
                                                              context) if right_operand.__class__ is ExpressionToken else OperandTokenTranslator.translate(
                right_operand, excel, context))

        if in_brackets:
            result.append(')')

        return ''.join(result)
