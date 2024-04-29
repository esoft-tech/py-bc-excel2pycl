from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import RoundControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class RoundControlConstructionTokenTranslator(AbstractTranslator[RoundControlConstructionToken]):
    @classmethod
    def translate(cls, token: RoundControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        number, num_digits = (
            ExpressionTokenTranslator.translate(token.number, excel, context),
            ExpressionTokenTranslator.translate(token.num_digits, excel, context),
        )

        return f"self._round({number}, {num_digits})"
