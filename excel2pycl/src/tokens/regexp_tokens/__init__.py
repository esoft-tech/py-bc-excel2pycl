import re
from excel2pycl.src.cell import Cell
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens.regexp_base_token import RegexpBaseToken


class CellIdentifierRangeToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\'|(\w*?))!)?((\$?([A-Z]+)(\$?(\d+))?:\$?\8(\$?(\d+))?)|(\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?\15)?))'
    last_match_regexp = r'([^\d$].*)?'
    value_range = [0, -1]

    def __init__(self, *args, **kwargs):
        self._range = None, None
        super().__init__(*args, *kwargs)

    def __str__(self):
        return f'<{self.__class__.__name__}>({self.range})'

    @property
    def range(self) -> (Cell, Cell,):
        if self._range[0] is None:
            self._range = Cell(title=self.value[3] or self.value[4] or self.in_cell.title, column=self.value[7] or self.value[13],
                               row=self.value[9] or self.value[15]), Cell(title=self.value[3] or self.value[4] or self.in_cell.title,
                                                                          column=self.value[7] or self.value[16],
                                                                          row=self.value[11] or self.value[15])
        return self._range


class MatrixOfCellIdentifiersToken(RegexpBaseToken):
    # TODO Consider the possibility of a matrix like A:A
    regexp = r'((\'(.*?)\'|(\w*?))!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'
    last_match_regexp = r'([^\d].*)?'
    value_range = [0, -1]

    def __init__(self, *args, **kwargs):
        self._matrix = None, None
        super().__init__(*args, *kwargs)

    def __str__(self):
        return f'<{self.__class__.__name__}>({self.matrix})'

    @property
    def matrix(self) -> (Cell, Cell,):
        if self._matrix[0] is None:
            self._matrix = Cell(title=self.value[3] or self.value[4] or self.in_cell.title, column=self.value[5],
                                row=self.value[7]), Cell(title=self.value[3] or self.value[4] or self.in_cell.title,
                                                         column=self.value[8], row=self.value[10])
        return self._matrix


class CellIdentifierToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\'|(\w*?))!)?\$?([A-Z]+)\$?(\d+)'
    last_match_regexp = r'([^\d]|[^:\d].*)?'
    value_range = [0, -1]

    def __init__(self, *args, **kwargs):
        self._cell = None
        super().__init__(*args, *kwargs)

    def __str__(self):
        return f'<{self.__class__.__name__}>({self.cell})'

    @property
    def cell(self) -> Cell:
        if self._cell is None:
            self._cell = Cell(title=self.value[3] or self.value[4] or self.in_cell.title, column=self.value[5],
                              row=self.value[6])
        return self._cell


class IfErrorKeywordToken(RegexpBaseToken):
    regexp = r'IFERROR'


class IfKeywordToken(RegexpBaseToken):
    regexp = r'(IF)\W'
    value_range = [1, 2]

    @classmethod
    def get(cls, expression: str, in_cell: Cell):
        result = re.findall(rf'^({cls.regexp})({cls.last_match_regexp})$', expression)

        if result:
            return cls(result[0][cls.value_range[0]:cls.value_range[1]], in_cell), result[0][-1]

        return None, expression


class MatchKeywordToken(RegexpBaseToken):
    regexp = r'MATCH'


class SumKeywordToken(RegexpBaseToken):
    regexp = r'SUM'


class MinKeywordToken(RegexpBaseToken):
    regexp = r'MIN'


class MaxKeywordToken(RegexpBaseToken):
    regexp = r'MAX'


class DayKeywordToken(RegexpBaseToken):
    regexp = r'DAY'


class MonthKeywordToken(RegexpBaseToken):
    regexp = r'MONTH'


class YearKeywordToken(RegexpBaseToken):
    regexp = r'YEAR'


class DateDiffKeywordToken(RegexpBaseToken):
    regexp = r'DATEDIF'


class DateKeywordToken(RegexpBaseToken):
    regexp = r'(DATE)\W'
    value_range = [1, 2]

    @classmethod
    def get(cls, expression: str, in_cell: Cell):
        result = re.findall(rf'^({cls.regexp})({cls.last_match_regexp})$', expression)

        if result:
            return cls(result[0][cls.value_range[0]:cls.value_range[1]], in_cell), result[0][-1]

        return None, expression


class VlookupKeywordToken(RegexpBaseToken):
    regexp = r'VLOOKUP'


class AverageKeywordToken(RegexpBaseToken):
    regexp = r'AVERAGE'


class RoundKeywordToken(RegexpBaseToken):
    regexp = r'ROUND'


class OrKeywordToken(RegexpBaseToken):
    regexp = r'OR'


class AndKeywordToken(RegexpBaseToken):
    regexp = r'AND'


# TODO добавить условие для локализации
class SeparatorToken(RegexpBaseToken):
    regexp = r';|,'


class EqOperatorToken(RegexpBaseToken):
    regexp = r'='


class NotEqOperatorToken(RegexpBaseToken):
    regexp = r'<>'


class GtOperatorToken(RegexpBaseToken):
    regexp = r'>'


class GtOrEqualOperatorToken(RegexpBaseToken):
    regexp = r'>='


class LtOperatorToken(RegexpBaseToken):
    regexp = r'<'


class LtOrEqualOperatorToken(RegexpBaseToken):
    regexp = r'<='


class PlusOperatorToken(RegexpBaseToken):
    regexp = r'\+'


class MinusOperatorToken(RegexpBaseToken):
    regexp = r'-'


class MultiplicationOperatorToken(RegexpBaseToken):
    regexp = r'\*'


class DivOperatorToken(RegexpBaseToken):
    regexp = r'/'


class AmpersandToken(RegexpBaseToken):
    regexp = r'&'


# TODO добавить условие для локализации
class LiteralToken(RegexpBaseToken):
    regexp = r'\"(.*?)\"|(\d+)((\.)(\d+))?(e(-?\d+))?|(TRUE(\(\))?)|(FALSE(\(\))?)'
    value_range = [0, -1]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, *kwargs)

        if self.value[2]:
            real_value = int(self.value[2])
            if self.value[5]:
                real_value += float(f'0.{self.value[5]}')
            if self.value[7]:
                # TODO in theory, the degree can be calculated using the expression
                real_value *= 10 ** int(self.value[7])
            real_value = str(real_value)
        elif self.value[1] or self.value[0] == '""':
            real_value = f'\'{self.value[1]}\''
        elif self.value[8]:
            real_value = 'True'
        elif self.value[10]:
            real_value = 'False'
        else:
            raise E2PyclParserException(f'Unknown literal value: {self.value}')

        self.value = real_value


class BracketStartToken(RegexpBaseToken):
    regexp = r'\('


class BracketFinishToken(RegexpBaseToken):
    regexp = r'\)'


class WhitespaceToken(RegexpBaseToken):
    regexp = r'[\s\n\t]+'
