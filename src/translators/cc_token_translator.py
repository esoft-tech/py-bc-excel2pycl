from src.context import Context
from src.excel import Excel
from src.tokens import IfControlConstructionToken, ControlConstructionToken, SumIfControlConstructionToken, \
    SumControlConstructionToken, AverageControlConstructionToken, VlookupControlConstructionToken
from src.translators.abstract_translator import AbstractTranslator


class ControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: ControlConstructionToken, excel: Excel, context: Context) -> str:
        from src.translators.if_cc_token_translator import IfControlConstructionTokenTranslator
        from src.translators.sum_if_cc_token_translator import SumIfControlConstructionTokenTranslator
        from src.translators.sum_cc_token_translator import SumControlConstructionTokenTranslator
        from src.translators.average_cc_token_translator import AverageControlConstructionTokenTranslator
        from src.translators.vlookup_cc_token_translator import VlookupControlConstructionTokenTranslator

        translate_functions = {
            IfControlConstructionToken.__name__: IfControlConstructionTokenTranslator.translate,
            SumIfControlConstructionToken.__name__: SumIfControlConstructionTokenTranslator.translate,
            SumControlConstructionToken.__name__: SumControlConstructionTokenTranslator.translate,
            AverageControlConstructionToken.__name__: AverageControlConstructionTokenTranslator.translate,
            VlookupControlConstructionToken.__name__: VlookupControlConstructionTokenTranslator.translate,
        }

        sub_token = token.control_construction
        translate_function = translate_functions.get(sub_token.__class__.__name__)
        if not translate_function:
            raise TypeError('Unknown control construction token')

        return translate_function(sub_token, excel, context)
