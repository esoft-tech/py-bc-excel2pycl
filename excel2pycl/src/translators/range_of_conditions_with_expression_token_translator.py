from excel2pycl.src.tokens import RangeOfConditionsWithExpressionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel


class RangeOfConditionsWithExpressionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: RangeOfConditionsWithExpressionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        from excel2pycl.src.translators.lambda_token_translator import LambdaTokenTranslator

        expression = ExpressionTokenTranslator.translate(token.expression, excel, context)
        condition = LambdaTokenTranslator.translate(token.condition_lambda, excel, context)

        return context.set_sub_cell(token.in_cell, f'{condition}, {expression}')
