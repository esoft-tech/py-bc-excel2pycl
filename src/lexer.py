import re


class Token:
    regexp = r''

    def __init__(self, value):
        self.value = value

    def __repr__(self):
        return f'{self.__class__}(`{self.value}`)'

    @classmethod
    def get(cls, expression: str):
        try:
            result = re.findall(rf'^({cls.regexp})(.*)', expression)
        except:
            print(cls, cls.regexp, expression)

        if result:
            return cls(result[0][0]), result[0][-1]

        return None, None


class MatrixOfCellIdentifiersToken(Token):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'


class SetOfCellIdentifiersToken(Token):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)(\$?(\d+))?:\$?([A-Z]+)(\$?(\d+))?'


class CellIdentifierToken(Token):
    regexp = r'((\'(.*?)\')!)?\$?([A-Z]+)\$?(\d+)'


class KeywordToken(Token):
    regexp = r'IF|SUM|SUMIF|VLOOKUP|AVERAGE'


# TODO добавить условие для локализации
class SeparatorToken(Token):
    regexp = r';|,'


class OperatorToken(Token):
    regexp = r'\+|-|/|\*|=|!=|>|>=|<|<=|!'


# TODO добавить условие для локализации
class LiteralToken(Token):
    regexp = r'\"(.*?)\"|(\d+)((,|.)(\d+))?(e(-?\d+))?|TRUE\(\)|FALSE\(\)'


class BracketStartToken(Token):
    regexp = r'\('


class BracketFinishToken(Token):
    regexp = r'\)'


class SpaceToken(Token):
    regexp = r'\s+'


class UndefinedToken(Token):
    @classmethod
    def get(cls, expression: str):
        raise Exception('Undefined token', expression)


class Lexer:
    TOKENS = [
        MatrixOfCellIdentifiersToken,
        SetOfCellIdentifiersToken,
        CellIdentifierToken,
        KeywordToken,
        SeparatorToken,
        OperatorToken,
        LiteralToken,
        BracketStartToken,
        BracketFinishToken,
        SpaceToken,
        UndefinedToken,
    ]

    @classmethod
    def parse(cls, expression):
        tokens = []
        while expression:
            for token_class in cls.TOKENS:
                token, sub_expression = token_class.get(expression)
                if token:
                    tokens.append(token)
                    expression = sub_expression
                    break

        return tokens
