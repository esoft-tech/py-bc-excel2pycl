from excel2pycl.src.tokens.composite_base_token import CompositeBaseToken

CLS = "cls"


class RecursiveCompositeBaseToken(CompositeBaseToken):
    @classmethod
    def get_token_sets(cls) -> list:
        if not cls._PROCESSED:
            cls._TOKEN_SETS = [[cls if token == CLS else token for token in tokens] for tokens in cls._TOKEN_SETS]
            cls._PROCESSED = True

        return cls._TOKEN_SETS
