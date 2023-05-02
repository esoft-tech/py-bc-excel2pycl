from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IfControlConstructionToken, ControlConstructionToken, SumIfControlConstructionToken, \
    SumControlConstructionToken, AverageControlConstructionToken, VlookupControlConstructionToken, \
    RoundControlConstructionToken, OrControlConstructionToken, AndControlConstructionToken, \
        MinControlConstructionToken, MaxControlConstructionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class ControlConstructionTokenTranslator(AbstractTranslator):
    @classmethod
    def translate(cls, token: ControlConstructionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.if_cc_token_translator import IfControlConstructionTokenTranslator
        from excel2pycl.src.translators.sum_if_cc_token_translator import SumIfControlConstructionTokenTranslator
        from excel2pycl.src.translators.sum_cc_token_translator import SumControlConstructionTokenTranslator
        from excel2pycl.src.translators.average_cc_token_translator import AverageControlConstructionTokenTranslator
        from excel2pycl.src.translators.vlookup_cc_token_translator import VlookupControlConstructionTokenTranslator
        from excel2pycl.src.translators.round_cc_token_translator import RoundControlConstructionTokenTranslator
        from excel2pycl.src.translators.or_cc_token_translator import OrControlConstructionTokenTranslator
        from excel2pycl.src.translators.and_cc_token_translator import AndControlConstructionTokenTranslator
        from excel2pycl.src.translators.min_cc_token_translator import MinControlConstructionTokenTranslator
        from excel2pycl.src.translators.max_cc_token_translator import MaxControlConstructionTokenTranslator

        translate_functions = {
            IfControlConstructionToken.__name__: IfControlConstructionTokenTranslator.translate,
            SumIfControlConstructionToken.__name__: SumIfControlConstructionTokenTranslator.translate,
            SumControlConstructionToken.__name__: SumControlConstructionTokenTranslator.translate,
            AverageControlConstructionToken.__name__: AverageControlConstructionTokenTranslator.translate,
            VlookupControlConstructionToken.__name__: VlookupControlConstructionTokenTranslator.translate,
            RoundControlConstructionToken.__name__: RoundControlConstructionTokenTranslator.translate,
            OrControlConstructionToken.__name__: OrControlConstructionTokenTranslator.translate,
            AndControlConstructionToken.__name__: AndControlConstructionTokenTranslator.translate,
            MinControlConstructionToken.__name__: MinControlConstructionTokenTranslator.translate,
            MaxControlConstructionToken.__name__: MaxControlConstructionTokenTranslator.translate
        }

        sub_token = token.control_construction
        translate_function = translate_functions.get(sub_token.__class__.__name__)
        if not translate_function:
            raise TypeError('Unknown control construction token')

        return translate_function(sub_token, excel, context)
