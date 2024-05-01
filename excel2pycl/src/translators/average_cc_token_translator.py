from typing import cast

from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import AverageControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator
from excel2pycl.src.utilities.helper import get_flatten_numeric_list


class AverageControlConstructionTokenTranslator(AbstractTranslator[AverageControlConstructionToken]):
    @classmethod
    def translate(cls, token: AverageControlConstructionToken, excel: Excel, context: Context) -> str:
        cls.check_token_type(token, AverageControlConstructionToken)
        token = cast(AverageControlConstructionToken, token)

        flatten_numeric_list = get_flatten_numeric_list(token, excel, context)

        return f"self._average({flatten_numeric_list})"
