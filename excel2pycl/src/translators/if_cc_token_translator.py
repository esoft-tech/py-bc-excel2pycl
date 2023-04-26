from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IfControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IfControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: IfControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        condition = ExpressionTokenTranslator.translate(token.condition, excel, context)
        when_true = ExpressionTokenTranslator.translate(token.when_true, excel, context)
        when_false = ExpressionTokenTranslator.translate(token.when_false, excel, context)

        return f'(({when_true}) if ({condition}) else ({when_false}))'
