from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import RightControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class RightControlConstructionTokenTranslator(AbstractTranslator[RightControlConstructionToken]):
    @classmethod
    def translate(cls, token: RightControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        text = ExpressionTokenTranslator.translate(token.text, excel, context)
        num_chars = ExpressionTokenTranslator.translate(token.num_chars, excel, context) if token.num_chars else None

        return f"self._right({text}, {num_chars})"
