from src.context import Context
from src.excel import Excel
from src.tokens import SumControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class SumControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: SumControlConstructionToken, excel: Excel, context: Context) -> str:
        from src.translators.expression_token_translator import ExpressionTokenTranslator
        flatten_list = context.set_sub_cell(token.in_cell, 'self._flatten_list([' + ','.join(
            [ExpressionTokenTranslator.translate(i, excel, context) for i in token.expressions]) + '])')
        flatten_numeric_list = context.set_sub_cell(token.in_cell, f'self._only_numeric_list({flatten_list})')
        return context.set_sub_cell(token.in_cell, f'self._sum({flatten_numeric_list})')
