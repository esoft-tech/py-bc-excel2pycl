from src.tokens.regexp_base_token import RegexpBaseToken


class MatrixOfCellIdentifiersToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'


class SetOfCellIdentifiersToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'


class CellIdentifierToken(RegexpBaseToken):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)\$?(\d+)'


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
    regexp = r'\"(.*?)\"|(\d+)((,|.)(\d+))?(e(-?\d+))?|TRUE\(\)|FALSE\(\)'


class BracketStartToken(RegexpBaseToken):
    regexp = r'\('


class BracketFinishToken(RegexpBaseToken):
    regexp = r'\)'


class WhitespaceToken(RegexpBaseToken):
    regexp = r'[\s\n\t]+'
