from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IfErrorControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IfErrorControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: IfErrorControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        # Если ошибка в ячейке, будет ли Exception, или же в ячейке будет N/A?
        try:
            return ExpressionTokenTranslator.translate(token.condition,
                                                       excel,
                                                       context)
        except Exception:
            return ExpressionTokenTranslator.translate(token.when_error, excel, context)
