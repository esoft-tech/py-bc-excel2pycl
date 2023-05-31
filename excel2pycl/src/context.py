from typing import Dict

from excel2pycl.src.cell import Cell


class Context:
    """
    Will be storing cell translation map.
    """

    def __init__(self):
        self._cell_translations = {}
        self._sub_cell_translations: Dict[str, list] = {}
        self._titles: Dict[str, int] = {}

    @property
    def __class_template(self) -> str:
        # TODO можно сделать кэш ячеек просчитанных
        return '''import datetime
from dateutil.relativedelta import relativedelta
from math import trunc
from typing import Dict, Literal
import calendar

class ExcelInPython:
    def __init__(self, arguments: list = None):
        if arguments is None:
            arguments = []
        self._arguments = {{}}
        self.set_arguments(arguments)
        self._titles = {titles}

    def set_arguments(self, arguments: list):
        self._arguments = {{
            **self._arguments,
            **{{i['uid']: i['value'] for i in arguments}}
        }}

    def get_titles(self) -> dict:
        return self._titles

    class EmptyCell(int):
        def __eq__(self, other):
            empty_cell_equal_values = ['', 0, None, False]
            if other in empty_cell_equal_values:
                return True
            return False

    def _flatten_list(self, subject: list) -> list:
        result = []
        for i in subject:
            if type(i) == list:
                result = result + self._flatten_list(i)
            else:
                result.append(i)

        return result

    def _find_error_in_list(self, flatten_list: list):
        for err_value in filter(lambda cell: cell in ['#NUM!', '#DIV/0!',
                                                      '#N/A', '#NAME?', ' #NULL!',
                                                      '#REF!', '#VALUE!'], flatten_list):
            return err_value

    @staticmethod
    def _only_numeric_list(flatten_list: list):
        return [i for i in flatten_list if type(i) in [float, int]]

    def _sum(self, flatten_list: list):
        return sum(self._only_numeric_list(flatten_list))

    def _average(self, flatten_list: list):
        return self._sum(flatten_list) / len(self._only_numeric_list(flatten_list))

    def _vlookup(self, lookup_value, table_array: list, col_index_num: int, range_lookup: bool = False):
        # TODO search optimization needed
        if not isinstance(range_lookup, (bool, int)):
            return '#ERROR!'

        if range_lookup:
            table_array.sort(key=lambda row: row[0])
        
        lookup_value_type = int if isinstance(lookup_value, self.EmptyCell) else type(lookup_value)

        if range_lookup:
            minimum = '#N/A'
            for row in table_array:
                if isinstance(row[0], self.EmptyCell) or not isinstance(row[0], lookup_value_type):
                    continue
                if row[0] <= lookup_value:
                    minimum = row[col_index_num - 1]
                else:
                    return minimum
        else:
            for row in table_array:
                if isinstance(row[0], self.EmptyCell) or not isinstance(row[0], lookup_value_type):
                    continue
                if row[0] == lookup_value:
                    return row[col_index_num - 1]

        return '#N/A'

    def _sum_if(self, range_: list, criteria: callable, sum_range: list = None):
        result = 0
        range_, sum_range = self._flatten_list(range_), self._flatten_list(sum_range)
        for i in range(len(range_)):
            if i < len(sum_range) and criteria(range_[i]):
                result += sum_range[i] or 0

        return result

    def _round(self, number: float, num_digits: int):
        return round(number, int(num_digits))
        
    def _date(self, year: int, month: int, day: int):
        match year:
            case year if 0 <= year <= 1899:
                year += 1900
            case year if year < 0 or year > 9999:
                return '#NUM!'

        result_date = datetime.datetime(year, 1, 1)

        result_date += relativedelta(months=month - 1)

        days_in_current_month = calendar.monthrange(result_date.year, result_date.month)[1]
        if abs(day) > days_in_current_month:
            while abs(day) > days_in_current_month:
                result_date += relativedelta(months=1 if (day > 0) else (-1))
                day += (-days_in_current_month) if day > 0 else days_in_current_month
                days_in_current_month = calendar.monthrange(result_date.year, result_date.month)[1]

        result_date += relativedelta(days=day - 1 if (day >= -1) else day - 2)

        return result_date

    def _datedif(self, date_start: datetime.datetime, date_end: datetime.datetime,
                 mode: Literal['Y', 'M', 'D', 'MD', 'YM', 'YD']):
        if (not isinstance(date_start, datetime.datetime) or not isinstance(date_end, datetime.datetime)):
            return "#VALUE!"
        if date_start > date_end:
            return "#NUM!"
        match mode:
            case 'Y':
                return (date_end - date_start).days // (366 if calendar.isleap(date_start.year) and
                                                        date_start.month <= 2 else 365)
            case 'M':
                result = 12 * (date_end.year - date_start.year) + (date_end.month - date_start.month)
                if date_start.day > date_end.day:
                    return result - 1
                return result
            case 'D':
                return (date_end - date_start).days
            case 'MD':
                if (date_end.day >= date_start.day):
                    return date_end.day - date_start.day
                else:
                    prev_month_date = datetime.datetime(date_end.year, date_end.month, 1) - datetime.timedelta(days=1)
                    return calendar.monthrange(prev_month_date.year, prev_month_date.month)[1] - (
                        date_start.day - date_end.day)
            case 'YM':
                return (12 if date_start.month > date_end.month and date_end.year > date_start.year else 0) \
                    + (date_end.month - date_start.month) \
                    + (-1 if date_start.day > date_end.day else 0)
            case 'YD':
                return (date_end - date_start).days % (366 if calendar.isleap(date_start.year) and
                                                       date_start.month <= 2 else 365)
            case _:
                return "#NUM!"

    def _eomonth(self, start_date: datetime.datetime, months: float | int):
        # Note: If months is not an integer, it is truncated.
        if not isinstance(start_date, datetime.datetime):
            return '#NUM!'
        result_date = start_date + relativedelta(months=trunc(months))
        last_day_num = calendar.monthrange(result_date.year, result_date.month)[1]
        return datetime.datetime(result_date.year, result_date.month, last_day_num)

    def _edate(self, start_date: datetime.datetime, months: float):
        if not isinstance(start_date, datetime.datetime):
            return '#VALUE!'
        if not isinstance(months, (int, float)):
            return '#VALUE!'
        return start_date + relativedelta(months=trunc(months))

    def _or(self, flatten_list: list):
        return any(flatten_list)

    def _and(self, flatten_list: list):
        return all(flatten_list)

    def _min(self, flatten_list: list):
        err_value = self._find_error_in_list(flatten_list)
        if err_value:
            return err_value

        return min(self._only_numeric_list(flatten_list))

    def _max(self, flatten_list: list):
        err_value = self._find_error_in_list(flatten_list)
        if err_value:
            return err_value

        return max(self._only_numeric_list(flatten_list))

    def _day(self, date: datetime.datetime):
        return date.day

    def _month(self, date: datetime.datetime):
        return date.month

    def _year(self, date: datetime.datetime):
        return date.year

    def _iferror(self, condition_function, when_error):
        try:
            cell = condition_function()
            if self._find_error_in_list([cell]):
                return when_error
            else:
                return cell
        except ZeroDivisionError:
            return when_error

    def _cell_preprocessor(self, cell_uid: str):
        return self._arguments.get(cell_uid, self.__dict__.get(cell_uid, self.__class__.__dict__[cell_uid])(self))

    def exec_function_in(self, cell_uid: str):
        return self._cell_preprocessor(cell_uid)

{functions}
'''

    @property
    def __function_template(self) -> str:
        return '''    def {name}(self):
        return {code}'''

    def __build_function(self, name: str, code: str) -> str:
        return self.__function_template.format(name=name, code=code)

    def __build_functions(self, functions: dict) -> str:
        return '\n\n'.join([self.__build_function(name, code) for name, code in functions.items()])

    def __build_class(self, functions: dict) -> str:
        return self.__class_template.format(functions=self.__build_functions(functions), titles=self._titles)

    @staticmethod
    def _get_cell_function_name(cell: Cell) -> str:
        # TODO Придумать как сделать это частью cell
        return cell.uid

    @classmethod
    def _get_sub_cell_function_name(cls, cell: Cell = None, sub_number: int = None, cell_prefix: str = None) -> str:
        return f'{cell_prefix or cls._get_cell_function_name(cell)}_{sub_number}'

    @staticmethod
    def _get_cell_with_cell_preprocessor(cell_function_name: str) -> str:
        return f"self._cell_preprocessor('{cell_function_name}')"

    def get_cell(self, cell: Cell) -> str or None:
        return self._get_cell_with_cell_preprocessor(
            self._get_cell_function_name(cell)) if cell.uid in self._cell_translations else None

    def set_cell(self, cell: Cell, code: str) -> str:
        self._cell_translations[self._get_cell_function_name(cell)] = code
        return self.get_cell(cell)

    def set_sub_cell(self, cell: Cell, code: str) -> str:
        # TODO check if sub expression exists
        cell_function_name = self._get_cell_function_name(cell)
        if not self._sub_cell_translations.get(cell_function_name):
            self._sub_cell_translations[self._get_cell_function_name(cell)] = []
        if code in self._sub_cell_translations[cell_function_name]:
            sub_number = self._sub_cell_translations[cell_function_name].index(code)
        else:
            self._sub_cell_translations[cell_function_name].append(code)
            sub_number = len(self._sub_cell_translations[cell_function_name]) - 1

        return self._get_cell_with_cell_preprocessor(self._get_sub_cell_function_name(cell=cell, sub_number=sub_number))

    def _get_divided_sub_cell_translations(self) -> dict:
        result = {}
        for cell_prefix, sub_cell_expressions in self._sub_cell_translations.items():
            for sub_cell_expression, sub_cell_index in zip(sub_cell_expressions, range(len(sub_cell_expressions))):
                result[self._get_sub_cell_function_name(cell_prefix=cell_prefix,
                                                        sub_number=sub_cell_index)] = sub_cell_expression

        return result

    def build_class(self) -> str:
        summary_functions = self._cell_translations.copy()
        summary_functions.update(self._get_divided_sub_cell_translations())

        return self.__build_class(summary_functions)
