import re

from excel2pycl.src.context import Context
from excel2pycl.src.excel import Excel
from excel2pycl.src.tokens import LambdaToken, PatternToken
from excel2pycl.src.translators.abstract_translator import AbstractTranslator


class LambdaTokenTranslator(AbstractTranslator[LambdaToken]):
    @classmethod
    def translate(cls, token: LambdaToken, excel: Excel, context: Context) -> str:
        from excel2pycl.src.translators.expression_token_translator import ExpressionTokenTranslator

        literal, expression = (
            token.literal,
            ExpressionTokenTranslator.translate(
                token.expression,
                excel,
                context,
            )
            if token.expression
            else None,
        )

        condition_symbol = "=="
        condition_value = literal

        if literal:
            if isinstance(literal, str):
                parsed_literal = re.findall(r"^\'(>=|<=|>|<|<>)((\d+)((\.)(\d+))?(e(-?\d+))?)?\'$", literal)
                if parsed_literal:
                    parsed_literal = parsed_literal[0]
                    if parsed_literal[0]:
                        condition_symbol = parsed_literal[0]
                        if condition_symbol == "<>":
                            condition_symbol = "!="

                    condition_value = parsed_literal[1] if parsed_literal[1] else expression

            else:
                if expression:
                    condition_value = expression
        else:
            condition_value = expression

        # ToDo: Если судить по функции получения expression, то этот if никогда не выполнится
        if getattr(getattr(token.expression, "left_operand", None), "value", None) and isinstance(
            token.expression.left_operand.value[0],
            PatternToken,
        ):
            return context.set_sub_cell(
                token.in_cell,
                f"lambda x: re.match({condition_value}, str(x))",
            )

        return context.set_sub_cell(
            token.in_cell,
            f"lambda x: "
            f"self._parse_date_obj(x){condition_symbol}self._parse_date_obj({condition_value}) "
            f"if self._parse_date_obj({condition_value}) "
            f"else str(x).lower(){condition_symbol}str({condition_value}).lower() "
            f"if isinstance({condition_value}, str) "
            f"else x{condition_symbol}{condition_value}",
        )
