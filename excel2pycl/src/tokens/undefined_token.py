from excel2pycl.src.tokens.base_token import BaseToken


class UndefinedToken(BaseToken):
    @classmethod
    def get(cls, expression: str, *args, **kwargs):
        raise Exception('Undefined token', expression, args)
