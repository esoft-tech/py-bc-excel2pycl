from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import EntryPointToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class EntryPointTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: EntryPointToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        if token is None:
            return 0
        return ExpressionTokenTranslator.translate(token.expression, excel, context)
