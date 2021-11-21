from src.tokens.base_token import BaseToken


class UndefinedToken(BaseToken):
    @classmethod
    def get(cls, expression: str):
        raise Exception('Undefined token', expression)
