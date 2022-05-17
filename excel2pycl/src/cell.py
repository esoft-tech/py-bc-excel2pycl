from dataclasses import dataclass


@dataclass
class Cell:
    title: int or str
    column: int or str
    row: int or str = None
    value: str or int or float or bool or None = None
    _handled_identifiers: bool = False

    def has_handled_identifiers(self):
        return self._handled_identifiers

    @property
    def uid(self) -> str:
        if not self._handled_identifiers and not (type(self.title) is int and type(self.column) is int and type(self.row) is int):
            raise Exception('Cell identifier must be integer when when used to get uid')

        return f"_{'_'.join([str(i) if i is not None else 'any' for i in [self.title, self.column, self.row]])}"

    def __str__(self):
        return f'<{self.__class__.__name__}>(`{self.title}`, `{self.column}`, {self.row})'

    def __repr__(self):
        return self.__str__()

    def to_dict(self) -> dict:
        return {'uid': self.uid, 'title': self.title, 'column': self.column, 'row': self.row, 'value': self.value}
