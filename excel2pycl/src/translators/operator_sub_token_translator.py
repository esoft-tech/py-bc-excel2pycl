from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import RegexpBaseToken, NotEqOperatorToken, EqOperatorToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class OperatorSubTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: RegexpBaseToken, excel: Excel, context: Context) -> str:
        operator = token.value[0]
        if token.__class__ == NotEqOperatorToken:
            operator = '!='
        elif token.__class__ == EqOperatorToken:
            operator = '=='

        return operator
