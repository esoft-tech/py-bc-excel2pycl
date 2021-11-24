from src.cell import Cell
from src.tokens.regexp_base_token import RegexpBaseToken


class CellIdentifierRangeToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?((\$?([A-Z]+)(\$?(\d+))?:\$?\7(\$?(\d+))?)|(\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?\15)?))'
    last_match_regexp = r'([^\d].*)?'
    value_range = [0, -1]

    def __init__(self, *args, **kwargs):
        self._range = None, None
        super().__init__(*args, *kwargs)

    def __str__(self):
        return f'<{self.__class__.__name__}>({self.range})'

    @property
    def range(self):
        if self._range[0] is None:
            self._range = Cell(title=self.value[3], column=self.value[6] or self.value[12],
                               row=self.value[8] or self.value[14]), Cell(title=self.value[3],
                                                                          column=self.value[6] or self.value[15],
                                                                          row=self.value[10] or self.value[14])
        return self._range


class MatrixOfCellIdentifiersToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'
    last_match_regexp = r'([^\d].*)?'
    value_range = [0, -1]

    def __init__(self, *args, **kwargs):
        self._matrix = None, None
        super().__init__(*args, *kwargs)

    def __str__(self):
        return f'<{self.__class__.__name__}>({self.matrix})'

    @property
    def matrix(self):
        if self._matrix[0] is None:
            self._matrix = Cell(title=self.value[3], column=self.value[4], row=self.value[6]), Cell(title=self.value[3],
                                                                                                    column=self.value[
                                                                                                        7],
                                                                                                    row=self.value[9])
        return self._matrix


class CellIdentifierToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)\$?(\d+)'
    last_match_regexp = r'([^\d]|[^:\d].*)?'
    value_range = [0, -1]

    def __init__(self, *args, **kwargs):
        self._cell = None
        super().__init__(*args, *kwargs)

    def __str__(self):
        return f'<{self.__class__.__name__}>({self.cell})'

    @property
    def cell(self):
        if self._cell is None:
            self._cell = Cell(title=self.value[3], column=self.value[4], row=self.value[5])
        return self._cell


class IfKeywordToken(RegexpBaseToken):
    regexp = r'IF'


class SumKeywordToken(RegexpBaseToken):
    regexp = r'SUM'


class SumIfKeywordToken(RegexpBaseToken):
    regexp = r'SUMIF'


class VlookupKeywordToken(RegexpBaseToken):
    regexp = r'VLOOKUP'


class AverageKeywordToken(RegexpBaseToken):
    regexp = r'AVERAGE'


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


class AndLambdaToken(RegexpBaseToken):
    regexp = r'&'


# TODO добавить условие для локализации
class LiteralToken(RegexpBaseToken):
    regexp = r'\"(.*?)\"|(\d+)((,|.)(\d+))?(e(-?\d+))?|(TRUE\(\))|(FALSE\(\))'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, *kwargs)

        if self.value[2]:
            real_value = int(self.value[2])
            if self.value[5]:
                real_value += float(f'0.{self.value[5]}')
            if self.value[7]:
                # TODO in theory, the degree can be calculated using the expression
                real_value *= 10**int(self.value[7])
        elif self.value[1]:
            real_value = self.value[1]
        elif self.value[8]:
            real_value = True
        elif self.value[9]:
            real_value = False
        else:
            raise Exception('Unknown literal value')

        self.value = real_value


class BracketStartToken(RegexpBaseToken):
    regexp = r'\('


class BracketFinishToken(RegexpBaseToken):
    regexp = r'\)'


class WhitespaceToken(RegexpBaseToken):
    regexp = r'[\s\n\t]+'
