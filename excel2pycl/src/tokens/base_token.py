from typing import Any, ClassVar

from excel2pycl.src.cell import Cell


class BaseToken:
    _SUBCLASSES: ClassVar[list] = []

    def __init__(self, value: Any, in_cell: Cell) -> None:  # noqa: ANN401
        self.value = value
        self._in_cell = in_cell

    def __str__(self) -> str:
        return f"<{self.__class__.__name__}>(value=`{self.value}`)"

    def __repr__(self) -> str:
        return self.__str__()

    @classmethod
    def get(cls, expression: str | list, in_cell: Cell) -> tuple["BaseToken | None", str | list]:
        raise NotImplementedError()

    @property
    def in_cell(self) -> Cell:
        return self._in_cell

    @classmethod
    def subclasses(cls) -> list:
        from excel2pycl.src.tokens.undefined_token import UndefinedToken

        if not cls._SUBCLASSES:
            subclasses = cls.__subclasses__()

            cls._SUBCLASSES = cls._remove_subclasses_lower_rank(subclasses)
            cls._SUBCLASSES.append(UndefinedToken)

        return cls._SUBCLASSES

    @classmethod
    def _remove_subclasses_lower_rank(cls, subclasses: list) -> list:
        from excel2pycl.src.tokens.undefined_token import UndefinedToken

        nested_classes = []
        for_difference_update = [UndefinedToken]
        for class_ in subclasses:
            if class_.__subclasses__():
                nested_classes += class_.subclasses()
                for_difference_update.append(class_)

        subclasses += nested_classes

        subclasses_set = set(subclasses)
        for subclass in for_difference_update:
            if subclass in subclasses_set:
                subclasses_set.discard(subclass)

        return list(subclasses)
