from src.context import Context
from src.excel import Excel
from src.tokens import IfControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class IfControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: IfControlConstructionToken, excel: Excel, context: Context) -> str:
        from src.translators.expression_token_translator import ExpressionTokenTranslator

        condition, when_true, when_false = ExpressionTokenTranslator.translate(token.condition, excel,
                                                                               context), ExpressionTokenTranslator.translate(
            token.when_true, excel, context), ExpressionTokenTranslator.translate(token.when_false, excel, context)

        return f'(({when_true}) if ({condition}) else ({when_false}))'
