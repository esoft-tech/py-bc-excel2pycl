class BaseToken:
    def __init__(self, value):
        self.value = value

    def __repr__(self):
        return f'<{self.__class__.__name__}>(value=`{self.value}`)'

    @classmethod
    def get(cls, expression: str or list):
        pass

    @classmethod
    def subclasses(cls) -> list:
        from src.tokens.composite_base_token import CompositeBaseToken
        from src.tokens.regexp_base_token import RegexpBaseToken

        subclasses = set(cls.__subclasses__() + CompositeBaseToken.subclasses() + RegexpBaseToken.subclasses())
        subclasses.difference_update({CompositeBaseToken, RegexpBaseToken})

        return list(subclasses)
