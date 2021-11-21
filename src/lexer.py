from src.tokens import RegexpBaseToken


class Lexer:
    TOKENS = RegexpBaseToken.subclasses()

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
