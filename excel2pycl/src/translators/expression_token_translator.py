from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ExpressionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ExpressionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: ExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.operand_token_translator import OperandTokenTranslator
        from excel2pycl.src.translators.operator_sub_token_translator import OperatorSubTokenTranslator

        left_brackets, right_brackets, operator, left_operand, right_operand = token.left_brackets, token.right_brackets, token.operator, token.left_operand, token.right_operand

        result = []

        if left_operand:
            if left_brackets:
                result.append('(')

            result.append(ExpressionTokenTranslator.translate(left_operand, excel,
                                                              context) if left_operand.__class__ is ExpressionToken else OperandTokenTranslator.translate(
                left_operand, excel, context))

            if left_brackets:
                result.append(')')

            if operator:
                result.append(OperatorSubTokenTranslator.translate(operator, excel, context))

        if right_operand:
            if right_brackets:
                result.append('(')

            if not left_operand and operator:
                result.append(OperatorSubTokenTranslator.translate(operator, excel, context))

            result.append(ExpressionTokenTranslator.translate(right_operand, excel,
                                                              context) if right_operand.__class__ is ExpressionToken else OperandTokenTranslator.translate(
                right_operand, excel, context))

            if right_brackets:
                result.append(')')

        return ''.join(result)
