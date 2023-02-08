from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import OrControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class OrControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: OrControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        flatten_list = context.set_sub_cell(token.in_cell, 'self._flatten_list([' + ','.join(
            [ExpressionTokenTranslator.translate(i, excel, context) for i in token.expressions]) + '])')

        return f'self._or({flatten_list})'
