from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import AmpersandToken, EqOperatorToken, NotEqOperatorToken, RegexpBaseToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class OperatorSubTokenTranslator(AbstractTranslator[RegexpBaseToken]):
    @classmethod
    def translate(cls, token: RegexpBaseToken, excel: Excel, context: Context) -> str:
        operator: str = token.value[0]
        if token.__class__ == NotEqOperatorToken:
            operator = "!="
        elif token.__class__ == EqOperatorToken:
            operator = "=="
        elif token.__class__ == AmpersandToken:
            operator = "+"

        return operator
