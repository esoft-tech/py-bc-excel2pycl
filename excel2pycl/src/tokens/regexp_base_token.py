import re
from typing import ClassVar

from excel2pycl.src.cell import Cell
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens.base_token import BaseToken


class RegexpBaseToken(BaseToken):
    regexp = r""
    last_match_regexp = r".*"
    value_range: ClassVar[int] = [0, 1]
    irreplaceable_value = None

    @classmethod
    def get(cls, expression: str | list, in_cell: Cell) -> tuple["BaseToken | None", str | list]:
        if isinstance(expression, list):
            raise E2PyclParserException("expression arg must be str, not list")

        result = re.findall(rf"^({cls.regexp})({cls.last_match_regexp})$", expression)

        if result:
            return cls(result[0][cls.value_range[0] : cls.value_range[1]], in_cell), result[0][-1]

        return None, expression


class KeywordRegexpBaseToken(RegexpBaseToken):
    """
    A subtype of regexp token containing the name of the function, allocated for their correct sorting
    in order to avoid dependence on the order of tokens in the library code
    """

    @classmethod
    def subclasses(cls) -> list:
        if not cls._SUBCLASSES:
            subclasses = cls.__subclasses__()
            subclasses_first_rank = cls._remove_subclasses_lower_rank(subclasses)
            cls._SUBCLASSES = sorted(subclasses_first_rank, key=lambda item: len(item.regexp), reverse=True)

        return cls._SUBCLASSES
