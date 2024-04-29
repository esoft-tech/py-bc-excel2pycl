from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.exceptions.e2pycl_translation_exception import E2PyclTranslationException
from excel2pycl.src.tokens import RangeOfCellIdentifierWithConditionToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class RangeOfCellIdentifierWithConditionTokenTranslator(AbstractTranslator[RangeOfCellIdentifierWithConditionToken]):
    @classmethod
    def translate(cls, token: RangeOfCellIdentifierWithConditionToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator
        from excel2pycl.src.translators.lambda_token_translator import LambdaTokenTranslator
        from excel2pycl.src.translators.matrix_of_cell_identifiers_token_translator import (
            MatrixOfCellIdentifiersTokenTranslator,
        )

        _range = MatrixOfCellIdentifiersTokenTranslator.translate(token.range, excel, context)
        if token.condition_lambda:
            condition = LambdaTokenTranslator.translate(token.condition_lambda, excel, context)
        else:
            if token.condition_expression is None:
                raise E2PyclTranslationException("Condition expression is None", token, context)

            condition = context.set_sub_cell(
                token.in_cell,
                f"lambda x: "
                f"self._parse_date_obj(x)==self._parse_date_obj({ExpressionTokenTranslator.translate(token.condition_expression, excel, context)}) "  # noqa: E501
                f"if self._parse_date_obj({ExpressionTokenTranslator.translate(token.condition_expression, excel, context)}) "  # noqa: E501
                f"else str(x).lower()==str({ExpressionTokenTranslator.translate(token.condition_expression, excel, context)}).lower() "  # noqa: E501
                f"if isinstance({ExpressionTokenTranslator.translate(token.condition_expression, excel, context)}, str) "  # noqa: E501
                f"else x=={ExpressionTokenTranslator.translate(token.condition_expression, excel, context)}",
            )

        return context.set_sub_cell(token.in_cell, f"{_range}, {condition}")
