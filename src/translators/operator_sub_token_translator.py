from src.context import Context
from src.excel import Excel
from src.tokens import RegexpBaseToken, NotEqOperatorToken, EqOperatorToken
from src.translators.abstract_translator import AbstractTranslator


class OperatorSubTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: RegexpBaseToken, excel: Excel, context: Context) -> str:
        operator = token.value[0]
        if token.__class__ == NotEqOperatorToken:
            operator = '!='
        elif token.__class__ == EqOperatorToken:
            operator = '=='

        return operator
