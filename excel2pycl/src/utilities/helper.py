from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import OperandToken
from excel2pycl.src.tokens.composite_base_token import CompositeBaseToken
from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator


def get_flatten_list(token: CompositeBaseToken, excel: Excel, context: Context) -> str:
    return context.set_sub_cell(token.in_cell, 'self._flatten_list([' + ','.join(
        [ExpressionTokenTranslator.translate(i, excel, context) for i in token.expressions]) + '])')


def get_flatten_numeric_list(token: CompositeBaseToken, excel: Excel, context: Context) -> str:
    return context.set_sub_cell(token.in_cell, f'self._only_numeric_list({get_flatten_list(token, excel, context)})')


def get_operand_source_token(token: OperandToken | None) -> CompositeBaseToken | None:
    """
    Ищет последний токен по вложенности
    """
    return get_operand_source_token(token.control_construction) if hasattr(token, 'control_construction') else token.__class__
