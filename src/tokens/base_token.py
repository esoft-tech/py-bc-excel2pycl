class BaseToken:
    _SUBCLASSES = []

    def __init__(self, value):
        self.value = value

    def __str__(self):
        return f'<{self.__class__.__name__}>(value=`{self.value}`)'

    def __repr__(self):
        return self.__str__()

    @classmethod
    def get(cls, expression: str or list):
        pass

    @classmethod
    def subclasses(cls) -> list:
        from src.tokens.undefined_token import UndefinedToken

        if not cls._SUBCLASSES:
            subclasses = cls.__subclasses__()
            nested_classes = []
            for_difference_update = [UndefinedToken]
            for class_ in subclasses:
                if class_.__subclasses__():
                    nested_classes += class_.subclasses()
                    for_difference_update.append(class_)

            subclasses += nested_classes
            subclasses = set(subclasses)
            subclasses.difference_update(set(for_difference_update))

            cls._SUBCLASSES = list(subclasses)
            cls._SUBCLASSES.append(UndefinedToken)

        return cls._SUBCLASSES
