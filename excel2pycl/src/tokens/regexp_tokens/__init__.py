from excel2pycl.src.cell import Cell
from excel2pycl.src.exceptions import E2PyclParserException
from excel2pycl.src.tokens.regexp_base_token import RegexpBaseToken, KeywordRegexpBaseToken


class MatrixOfCellIdentifiersToken(RegexpBaseToken):
    # TODO Consider the possibility of a matrix like A:A
    regexp = r'((\'([^!]*?)\'|(\w*?))!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'
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


class PatternToken(RegexpBaseToken):
    """
    Parses excel pattern system
    ? - stands for single symbol
    * - stands for sequence of symbols
    ~ - cancels pattern effect if placed before ? or * (cancels effect only for next symbol, but not for all)
    Would be useful to recognize argument in function e.g =COUNTIFS(A3:B3; "???le") or =COUNTIFS(A4:B7; "a*")
    """
    regexp = r'\"(.*(?<![~])[?*]+.*)\"'


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
            raise E2PyclParserException('Unknown literal value')

        self.value = real_value


class BracketStartToken(RegexpBaseToken):
    regexp = r'\('


class BracketFinishToken(RegexpBaseToken):
    regexp = r'\)'


class WhitespaceToken(RegexpBaseToken):
    regexp = r'[\s\n\t]+'


# TODO добавить условие для локализации
class SeparatorToken(RegexpBaseToken):
    regexp = r';|,|~'


class NotEqOperatorToken(RegexpBaseToken):
    regexp = r'<>'


class GtOrEqualOperatorToken(RegexpBaseToken):
    regexp = r'>='


class LtOrEqualOperatorToken(RegexpBaseToken):
    regexp = r'<='


class EqOperatorToken(RegexpBaseToken):
    regexp = r'='


class GtOperatorToken(RegexpBaseToken):
    regexp = r'>'


class LtOperatorToken(RegexpBaseToken):
    regexp = r'<'


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


class PercentToken(RegexpBaseToken):
    regexp = r'%'


class AddressKeywordToken(KeywordRegexpBaseToken):
    regexp = r'ADDRESS'


class AndKeywordToken(KeywordRegexpBaseToken):
    regexp = r'AND'


class AverageKeywordToken(KeywordRegexpBaseToken):
    regexp = r'AVERAGE'


class AverageIfSKeywordToken(KeywordRegexpBaseToken):
    regexp = r'AVERAGEIFS'


class ColumnKeywordToken(KeywordRegexpBaseToken):
    regexp = r'COLUMN'


class CountKeywordToken(KeywordRegexpBaseToken):
    regexp = r'COUNT'


class CountBlankKeywordToken(KeywordRegexpBaseToken):
    regexp = r'COUNTBLANK'


class CountIfSKeywordToken(KeywordRegexpBaseToken):
    regexp = r'COUNTIFS'


class DayKeywordToken(KeywordRegexpBaseToken):
    regexp = r'DAY'


class DateKeywordToken(KeywordRegexpBaseToken):
    regexp = r'DATE'


class DateDifKeywordToken(KeywordRegexpBaseToken):
    regexp = r'DATEDIF'


class EDateKeywordToken(KeywordRegexpBaseToken):
    regexp = r'EDATE'


class EoMonthKeywordToken(KeywordRegexpBaseToken):
    regexp = r'EOMONTH'


class IfKeywordToken(KeywordRegexpBaseToken):
    regexp = r'IF'


class IfErrorKeywordToken(KeywordRegexpBaseToken):
    regexp = r'IFERROR'


class IndexKeywordToken(KeywordRegexpBaseToken):
    regexp = r'INDEX'


class LeftKeywordToken(KeywordRegexpBaseToken):
    regexp = r'LEFT'


class MatchKeywordToken(KeywordRegexpBaseToken):
    regexp = r'MATCH'


class MaxKeywordToken(KeywordRegexpBaseToken):
    regexp = r'MAX'


class MidKeywordToken(KeywordRegexpBaseToken):
    regexp = r'MID'


class MinKeywordToken(KeywordRegexpBaseToken):
    regexp = r'MIN'


class MonthKeywordToken(KeywordRegexpBaseToken):
    regexp = r'MONTH'


class NetworkDaysKeywordToken(KeywordRegexpBaseToken):
    regexp = r'NETWORKDAYS'


class OrKeywordToken(KeywordRegexpBaseToken):
    regexp = r'OR'


class RightKeywordToken(KeywordRegexpBaseToken):
    regexp = r'RIGHT'


class RoundKeywordToken(KeywordRegexpBaseToken):
    regexp = r'ROUND'


class RoundUpKeywordToken(KeywordRegexpBaseToken):
    regexp = r'ROUNDUP'


class SearchKeywordToken(KeywordRegexpBaseToken):
    regexp = r'SEARCH'


class SumKeywordToken(KeywordRegexpBaseToken):
    regexp = r'SUM'


class SumIfKeywordToken(KeywordRegexpBaseToken):
    regexp = r'SUMIF'


class SumIfSKeywordToken(KeywordRegexpBaseToken):
    regexp = r'SUMIFS'


class TodayKeywordToken(KeywordRegexpBaseToken):
    regexp = r'TODAY'


class VlookupKeywordToken(KeywordRegexpBaseToken):
    regexp = r'VLOOKUP'


class XMatchKeywordToken(KeywordRegexpBaseToken):
    regexp = r'XMATCH'


class YearKeywordToken(KeywordRegexpBaseToken):
    regexp = r'YEAR'


class IfsKeywordToken(KeywordRegexpBaseToken):
    regexp = r'IFS'
