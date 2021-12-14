from abc import abstractmethod

from src.context import Context
from src.excel import Excel
from src.tokens import BaseToken


class AbstractTranslator:
    @classmethod
    @abstractmethod
    def translate(cls, token: BaseToken, excel: Excel, context: Context) -> str:
        pass

    @classmethod
    def translate_to_sub_cell(cls, token: BaseToken, excel: Excel, context: Context) -> str:
        return context.set_sub_cell(token.in_cell, cls.translate(token, excel, context))
