from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IfControlConstructionToken, ControlConstructionToken, SumIfControlConstructionToken, \
    SumControlConstructionToken, AverageControlConstructionToken, VlookupControlConstructionToken, \
    RoundControlConstructionToken, OrControlConstructionToken, AndControlConstructionToken, \
    MinControlConstructionToken, MaxControlConstructionToken, \
    YearControlConstructionToken, MonthControlConstructionToken, DayControlConstructionToken, \
    IfErrorControlConstructionToken, \
    DateControlConstructionToken, \
    DateDifControlConstructionToken, \
    EoMonthControlConstructionToken, \
    EDateControlConstructionToken, MatchControlConstructionToken, XMatchControlConstructionToken, \
    LeftControlConstructionToken, MidControlConstructionToken, RightControlConstructionToken, \
    CountBlankControlConstructionToken, \
    SearchControlConstructionToken, TodayControlConstructionToken, \
    AverageIfsControlConstructionToken, AddressControlConstructionToken, CountIfsControlConstructionToken, \
    NetworkDaysControlConstructionToken, CountControlConstructionToken
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
        from excel2pycl.src.translators.year_cc_token_translator import YearControlConstructionTokenTranslator
        from excel2pycl.src.translators.month_cc_token_translator import MonthControlConstructionTokenTranslator
        from excel2pycl.src.translators.day_cc_token_translator import DayControlConstructionTokenTranslator
        from excel2pycl.src.translators.iferror_cc_token_translator import IfErrorControlConstructionTokenTranslator
        from excel2pycl.src.translators.date_cc_token_translator import DateControlConstructionTokenTranslator
        from excel2pycl.src.translators.date_dif_cc_token_translator import DateDifControlConstructionTokenTranslator
        from excel2pycl.src.translators.eo_month_cc_token_translator import EoMonthControlConstructionTokenTranslator
        from excel2pycl.src.translators.e_date_cc_token_translator import EDateControlConstructionTokenTranslator
        from excel2pycl.src.translators.match_cc_token_translator import MatchControlConstructionTokenTranslator
        from excel2pycl.src.translators.xmatch_cc_token_translator import XMatchControlConstructionTokenTranslator
        from excel2pycl.src.translators.left_cc_token_translator import LeftControlConstructionTokenTranslator
        from excel2pycl.src.translators.mid_cc_token_translator import MidControlConstructionTokenTranslator
        from excel2pycl.src.translators.right_cc_token_translator import RightControlConstructionTokenTranslator
        from excel2pycl.src.translators.averageifs_cc_token_translator import \
            AverageIfsControlConstructionTokenTranslator
        from excel2pycl.src.translators.countblank_cc_token_translator import \
            CountBlankControlConstructionTokenTranslator
        from excel2pycl.src.translators.search_cc_token_translator import SearchControlConstructionTokenTranslator
        from excel2pycl.src.translators.today_cc_token_translator import TodayControlConstructionTokenTranslator
        from excel2pycl.src.translators.countifs_cc_token_translator import CountIfsControlConstructionTokenTranslator
        from excel2pycl.src.translators.address_cc_token_translator import AddressControlConstructionTokenTranslator
        from excel2pycl.src.translators.networkdays_cc_token_translator import \
            NetworkDaysControlConstructionTokenTranslator
        from excel2pycl.src.translators.count_cc_token_translator import CountControlConstructionTokenTranslator

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
            MaxControlConstructionToken.__name__: MaxControlConstructionTokenTranslator.translate,
            YearControlConstructionToken.__name__: YearControlConstructionTokenTranslator.translate,
            MonthControlConstructionToken.__name__: MonthControlConstructionTokenTranslator.translate,
            DayControlConstructionToken.__name__: DayControlConstructionTokenTranslator.translate,
            IfErrorControlConstructionToken.__name__: IfErrorControlConstructionTokenTranslator.translate,
            DateControlConstructionToken.__name__: DateControlConstructionTokenTranslator.translate,
            DateDifControlConstructionToken.__name__: DateDifControlConstructionTokenTranslator.translate,
            EoMonthControlConstructionToken.__name__: EoMonthControlConstructionTokenTranslator.translate,
            EDateControlConstructionToken.__name__: EDateControlConstructionTokenTranslator.translate,
            MatchControlConstructionToken.__name__: MatchControlConstructionTokenTranslator.translate,
            XMatchControlConstructionToken.__name__: XMatchControlConstructionTokenTranslator.translate,
            LeftControlConstructionToken.__name__: LeftControlConstructionTokenTranslator.translate,
            MidControlConstructionToken.__name__: MidControlConstructionTokenTranslator.translate,
            RightControlConstructionToken.__name__: RightControlConstructionTokenTranslator.translate,
            AverageIfsControlConstructionToken.__name__: AverageIfsControlConstructionTokenTranslator.translate,
            CountBlankControlConstructionToken.__name__: CountBlankControlConstructionTokenTranslator.translate,
            SearchControlConstructionToken.__name__: SearchControlConstructionTokenTranslator.translate,
            CountControlConstructionToken.__name__: CountControlConstructionTokenTranslator.translate,
            AddressControlConstructionToken.__name__: AddressControlConstructionTokenTranslator.translate,
            CountIfsControlConstructionToken.__name__: CountIfsControlConstructionTokenTranslator.translate,
            TodayControlConstructionToken.__name__: TodayControlConstructionTokenTranslator.translate,
            NetworkDaysControlConstructionToken.__name__: NetworkDaysControlConstructionTokenTranslator.translate
        }

        sub_token = token.control_construction
        translate_function = translate_functions.get(sub_token.__class__.__name__)
        if not translate_function:
            raise TypeError('Unknown control construction token')

        return translate_function(sub_token, excel, context)
