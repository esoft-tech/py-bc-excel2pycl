import re

from src.context import Context
from src.excel import Excel
from src.tokens import LambdaToken
from src.translators.abstract_translator import AbstractTranslator


class LambdaTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: LambdaToken, excel: Excel, context: Context) -> str:
        from src.translators import ExpressionTokenTranslator
        literal, expression = token.literal, ExpressionTokenTranslator.translate(token.expression, excel,
                                                                                 context) if token.expression else None

        condition_symbol = '=='
        condition_value = literal

        if literal:
            parsed_literal = re.findall(r'^\'(>=|<=|>|<|<>)((\d+)((\.)(\d+))?(e(-?\d+))?)?\'$', literal)
            if parsed_literal:
                parsed_literal = parsed_literal[0]
                if parsed_literal[0]:
                    condition_symbol = parsed_literal[0]
                    if condition_symbol == '<>':
                        condition_symbol = '!='

                if parsed_literal[1]:
                    condition_value = parsed_literal[1]
                else:
                    condition_value = expression
        else:
            condition_value = expression

        return context.set_sub_cell(token.in_cell, f'lambda x:x{condition_symbol}{condition_value}')
