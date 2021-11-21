from src.tokens.regexp_base_token import RegexpBaseToken


class MatrixOfCellIdentifiersToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'


class SetOfCellIdentifiersToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'


class CellIdentifierToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)\$?(\d+)'


class KeywordToken(RegexpBaseToken):
    regexp = r'IF|SUM|SUMIF|VLOOKUP|AVERAGE'


# TODO добавить условие для локализации
class SeparatorToken(RegexpBaseToken):
    regexp = r';|,'


class OperatorToken(RegexpBaseToken):
    regexp = r'\+|-|/|\*|=|!=|>|>=|<|<=|!'


# TODO добавить условие для локализации
class LiteralToken(RegexpBaseToken):
    regexp = r'\"(.*?)\"|(\d+)((,|.)(\d+))?(e(-?\d+))?|TRUE\(\)|FALSE\(\)'


class BracketStartToken(RegexpBaseToken):
    regexp = r'\('


class BracketFinishToken(RegexpBaseToken):
    regexp = r'\)'


class WhitespaceToken(RegexpBaseToken):
    regexp = r'[\s\n\t]+'
