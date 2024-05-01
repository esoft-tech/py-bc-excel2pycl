from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import IterableRangeOfCellIdentifierWithConditionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class IterableRangeOfCellIdentifierWithConditionTokenTranslator(
    AbstractTranslator[IterableRangeOfCellIdentifierWithConditionToken],
):
    @classmethod
    def translate(cls, token: IterableRangeOfCellIdentifierWithConditionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.range_of_cell_identifier_with_condition_token_translator import (
            RangeOfCellIdentifierWithConditionTokenTranslator,
        )

        return context.set_sub_cell(
            token.in_cell,
            "["
            + ", ".join(
                [
                    f"*{RangeOfCellIdentifierWithConditionTokenTranslator.translate(i, excel, context)}"
                    for i in token.range_of_cell_identifiers_with_conditions
                ],
            )
            + "]",
        )
