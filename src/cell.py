from dataclasses import dataclass


@dataclass
class Cell:
    title: int or str
    column: int or str
    row: int or None = None

    def __str__(self):
        return f'<{self.__class__.__name__}>(`{self.title}`, `{self.column}`, {self.row})'

    def __repr__(self):
        return self.__str__()
