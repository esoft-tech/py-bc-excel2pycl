from abc import abstractmethod

from src.cell import Cell
from src.context import Context
from src.excel import Excel


class AbstractTranslator:
    @classmethod
    @abstractmethod
    def translate(cls, cell: Cell, excel: Excel, context: Context) -> str:
        pass
