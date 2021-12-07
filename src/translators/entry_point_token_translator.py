from src.context import Context
from src.excel import Excel
from src.tokens import EntryPointToken
from src.translators.abstract_translator import AbstractTranslator


class EntryPointTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: EntryPointToken, excel: Excel, context: Context) -> str:
        from src.translators.expression_token_translator import ExpressionTokenTranslator

        return ExpressionTokenTranslator.translate(token.expression, excel, context)
