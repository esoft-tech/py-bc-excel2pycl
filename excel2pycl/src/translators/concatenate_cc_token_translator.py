from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import ConcatenateControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ConcatenateControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: ConcatenateControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import \
            ExpressionTokenTranslator

        expressions = [
            ExpressionTokenTranslator.translate(expression, excel, context) for expression in token.expressions]

        return '+'.join([f'str({expression})' for expression in expressions])
