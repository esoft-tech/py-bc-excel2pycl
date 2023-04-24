from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import MatchControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IfErrorControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: MatchControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        lookup_value, lookup_array, match_type = \
            ExpressionTokenTranslator.translate(token.lookup_value,
                                                excel,
                                                context),
        ExpressionTokenTranslator.translate(token.lookup_array,
                                            excel,
                                            context),
        ExpressionTokenTranslator.translate(token.match_type,
                                            excel,
                                            context)
