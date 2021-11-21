from src.tokens.undefined_token import UndefinedToken
from src.tokens.composite_base_token import CompositeBaseToken


class RecursiveCompositeBaseToken(CompositeBaseToken):
    @classmethod
    def subclasses(cls) -> list:
        return cls.__subclasses__() + [UndefinedToken]
