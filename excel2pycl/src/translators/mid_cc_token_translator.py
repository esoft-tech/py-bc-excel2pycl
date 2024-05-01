from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import MidControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class MidControlConstructionTokenTranslator(AbstractTranslator[MidControlConstructionToken]):
    @classmethod
    def translate(cls, token: MidControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        text = ExpressionTokenTranslator.translate(token.text, excel, context)
        start_num = ExpressionTokenTranslator.translate(token.start_num, excel, context)
        num_chars = ExpressionTokenTranslator.translate(token.num_chars, excel, context)

        return f"self._mid({text}, {start_num}, {num_chars})"
