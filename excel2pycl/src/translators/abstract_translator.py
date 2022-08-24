from abc import abstractmethod, ABC

from excel2pycl.src.cell import Cell
from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens.base_token import BaseToken


class AbstractTranslator(ABC):
    @classmethod
    @abstractmethod
    def translate(cls, subject: BaseToken or Cell, excel: Excel, context: Context) -> str:
        pass
