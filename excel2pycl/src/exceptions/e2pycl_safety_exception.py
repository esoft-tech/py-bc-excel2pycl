from excel2pycl.src.exceptions.e2pycl_parser_exception import E2PyclParserException


class E2PyclSafetyException(E2PyclParserException):
    def __init__(self, *args, **kwargs):
        self.suspicious_cells = kwargs.get('suspicious_cells', {})
        super().__init__('\n\t'.join([f"{k}: {v}" for k, v in self.suspicious_cells.items()]))
