from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens.base_token import BaseToken


class UndefinedToken(BaseToken):
    @classmethod
    def get(cls, expression: str, *args, **kwargs):
        raise E2PyclParserException('Undefined token', expression, args)
