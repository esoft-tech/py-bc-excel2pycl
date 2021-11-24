import re

from src.tokens.undefined_token import UndefinedToken
from src.tokens.base_token import BaseToken


class RegexpBaseToken(BaseToken):
    regexp = r''

    @classmethod
    def get(cls, expression: str):
        result = re.findall(rf'^({cls.regexp})(.*)', expression)

        if result:
            return cls(result[0][0]), result[0][-1]

        return None, expression
