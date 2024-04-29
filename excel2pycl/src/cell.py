from dataclasses import dataclass

from excel2pycl.src.exceptions import E2PyclCellException


@dataclass
class Cell:
    title: int | str
    column: int | str
    row: int | str | None = None
    value: str | int | float | bool | None = None
    _handled_identifiers: bool = False

    def has_handled_identifiers(self) -> bool:
        return self._handled_identifiers

    @property
    def uid(self) -> str:
        if not self._handled_identifiers and not (isinstance(self.column, int) and isinstance(self.row, int)):
            raise E2PyclCellException("Cell column and row identifiers must be integer when used to get uid")

        return f"_{'_'.join([str(i) if i is not None else 'any' for i in [self.title, self.column, self.row]])}"

    def __str__(self) -> str:
        return f"<{self.__class__.__name__}>(`{self.title}`, `{self.column}`, {self.row})"

    def __repr__(self) -> str:
        return self.__str__()

    def __hash__(self) -> int:
        return hash(frozenset([self.uid, self.value]))

    def to_dict(self) -> dict:
        return {"uid": self.uid, "title": self.title, "column": self.column, "row": self.row, "value": self.value}
