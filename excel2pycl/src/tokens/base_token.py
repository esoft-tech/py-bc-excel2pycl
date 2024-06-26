from excel2pycl.src.cell import Cell


class BaseToken:
    _SUBCLASSES = []

    def __init__(self, value, in_cell: Cell):
        self.value = value
        self._in_cell = in_cell

    def __str__(self):
        return f'<{self.__class__.__name__}>(value=`{self.value}`)'

    def __repr__(self):
        return self.__str__()

    @classmethod
    def get(cls, expression: str or list, in_cell: Cell):
        pass

    @property
    def in_cell(self) -> Cell:
        return self._in_cell

    @classmethod
    def subclasses(cls) -> list:
        from excel2pycl.src.tokens.undefined_token import UndefinedToken

        if not cls._SUBCLASSES:
            subclasses = cls.__subclasses__()
            subclasses_first_rank = cls._remove_subclasses_lower_rank(subclasses)

            cls._SUBCLASSES = list(subclasses_first_rank)
            cls._SUBCLASSES.append(UndefinedToken)

        return cls._SUBCLASSES

    def _remove_subclasses_lower_rank(subclasses: list) -> dict:
        from excel2pycl.src.tokens.undefined_token import UndefinedToken

        nested_classes = []
        for_difference_update = [UndefinedToken]
        for class_ in subclasses:
            if class_.__subclasses__():
                nested_classes += class_.subclasses()
                for_difference_update.append(class_)

        subclasses += nested_classes

        subclasses = dict.fromkeys(subclasses)
        for subclass in for_difference_update:
            if subclass in subclasses:
                del subclasses[subclass]

        return subclasses
