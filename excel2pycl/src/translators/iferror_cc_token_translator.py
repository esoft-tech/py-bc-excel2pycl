from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IfErrorControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IfErrorControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: IfErrorControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        condition = ExpressionTokenTranslator.translate(token.condition, excel, context)
        when_error = ExpressionTokenTranslator.translate(token.when_error, excel, context)

        # return context.set_sub_cell(token.in_cell, f'lambda x:x{condition}{when_error}')
        return f'self._iferror(lambda x: {condition}, {when_error})'
