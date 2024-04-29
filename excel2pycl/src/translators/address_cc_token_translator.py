from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import AddressControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator


class AddressControlConstructionTokenTranslator(AbstractTranslator[AddressControlConstructionToken]):
    @classmethod
    def translate(cls, token: AddressControlConstructionToken, excel: Excel, context: Context) -> str:
        cls.check_token_type(token, AddressControlConstructionToken)

        args = [ExpressionTokenTranslator.translate(expression, excel, context) for expression in token.expressions]
        row = ExpressionTokenTranslator.translate(token.row, excel, context)
        col = ExpressionTokenTranslator.translate(token.col, excel, context)

        return f"self._address({row}, {col}, *{args})"
