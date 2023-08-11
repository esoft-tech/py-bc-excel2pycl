from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import SearchControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class SearchControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: SearchControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        find_text = ExpressionTokenTranslator.translate(token.find_text, excel, context)
        within_text = ExpressionTokenTranslator.translate(token.within_text, excel, context)

        start_num = ExpressionTokenTranslator.translate(token.start_num, excel, context) if token.start_num else None

        return f'self._search({find_text}, {within_text}, {start_num})'
