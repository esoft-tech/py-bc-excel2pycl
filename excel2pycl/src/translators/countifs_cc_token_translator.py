from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import CountIfsControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.translators.range_of_cell_identifier_with_condition_token_translator import \
    RangeOfCellIdentifierWithConditionTokenTranslator


class CountIfsControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: CountIfsControlConstructionToken, excel: Excel, context: Context) -> str:
        translated_args = RangeOfCellIdentifierWithConditionTokenTranslator.translate(
            token.range_n_condition, excel, context
        )
        return f'self._countifs(*{translated_args})'
