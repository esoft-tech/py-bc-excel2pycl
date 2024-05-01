from typing import Generic, TypeVar

from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens.base_token import BaseToken

T = TypeVar("T")


class AbstractTranslator(Generic[T]):
    @classmethod
    def translate(cls, subject: T, excel: Excel, context: Context) -> str:
        raise NotImplementedError()

    @classmethod
    def check_token_type(cls, token: T, expected_type: type[BaseToken]) -> None:
        if not isinstance(token, expected_type):
            raise TypeError(f"token has wrong type: {type(token)}")
