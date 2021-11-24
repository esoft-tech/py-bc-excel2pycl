from src.tokens import EntryPointToken


class AstBuilder:
    @classmethod
    def parse(cls, expression: list):
        return EntryPointToken.get(expression)
