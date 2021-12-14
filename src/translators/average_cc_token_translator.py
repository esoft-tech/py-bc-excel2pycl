from src.context import Context
from src.excel import Excel
from src.tokens import AverageControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class AverageControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: AverageControlConstructionToken, excel: Excel, context: Context) -> str:
        from src.translators.expression_token_translator import ExpressionTokenTranslator
        flatten_list = context.set_sub_cell(token.in_cell, 'self._flatten_list([' + ','.join(
            [ExpressionTokenTranslator.translate(i, excel, context) for i in token.expressions]) + '])')
        flatten_list_sum = context.set_sub_cell(token.in_cell, f'sum({flatten_list})')

        return f'({flatten_list_sum}/len({flatten_list}))'
