import re

from src.cell import Cell
from src.tokens.base_token import BaseToken


class RegexpBaseToken(BaseToken):
    regexp = r''
    last_match_regexp = r'.*'
    value_range = [0, 1]
    irreplaceable_value = None

    @classmethod
    def get(cls, expression: str, in_cell: Cell):
        result = re.findall(rf'^({cls.regexp})({cls.last_match_regexp})$', expression)

        if result:
            return cls(result[0][cls.value_range[0]:cls.value_range[1]], in_cell), result[0][-1]

        return None, expression
