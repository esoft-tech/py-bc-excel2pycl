from abc import ABC
import calendar
import datetime
from typing import Dict, Literal


class AbstractExcelInPython(ABC):
    def __init__(self, arguments: list = None):
        if arguments is None:
            arguments = []
        self._arguments = {}
        self.set_arguments(arguments)
        self._titles = {}

    def set_arguments(self, arguments: list):
        self._arguments = {
            **self._arguments,
            **{i['uid']: i['value'] for i in arguments}
        }

    def get_titles(self) -> Dict[str, int]:
        return self._titles

    def _flatten_list(self, subject: list) -> list:
        result = []
        for i in subject:
            if type(i) == list:
                result = result + self._flatten_list(i)
            else:
                result.append(i)

        return result

    @staticmethod
    def _only_numeric_list(flatten_list: list):
        return [i for i in flatten_list if type(i) in [float, int]]

    def _sum(self, flatten_list: list):
        return sum(self._only_numeric_list(flatten_list))

    def _average(self, flatten_list: list):
        return self._sum(flatten_list) / len(self._only_numeric_list(flatten_list))

    def _vlookup(self, lookup_value, table_array: list, col_index_num: int, range_lookup: bool = False):
        # TODO add Range Lookup (https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)
        # TODO search optimization needed
        for row in table_array:
            if row[0] == lookup_value:
                return row[col_index_num - 1]

        return 1  # When the vlookup has not found anything similar to lookup_value, it returns 1 in the excel implementation for this function

    def _sum_if(self, range_: list, criteria: callable, sum_range: list = None):
        result = 0
        range_, sum_range = self._flatten_list(range_), self._flatten_list(sum_range)
        for i in range(len(range_)):
            if i < len(sum_range) and criteria(range_[i]):
                result += sum_range[i] or 0

        return result

    def _round(self, number: float, num_digits: int):
        return round(number, int(num_digits))

    def _datedif(self, date_start: datetime.datetime, date_end: datetime.datetime,
                 mode: Literal['Y', 'M', 'D', 'MD', 'YM', 'YD']):
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
                end = datetime.datetime(date_start.year + 1, date_end.month, date_end.day)
                return (end - date_start).days
            case _:
                raise Exception('Invalid mode value for DATEDIF')

    def _or(self, flatten_list: list):
        return any(flatten_list)

    def _and(self, flatten_list: list):
        return all(flatten_list)

    def _cell_preprocessor(self, cell_uid: str):
        return self._arguments.get(cell_uid, self.__dict__.get(cell_uid, self.__class__.__dict__[cell_uid])(self))

    def exec_function_in(self, cell_uid: str):
        return self._cell_preprocessor(cell_uid)

    class EmptyCell(int):
        def __eq__(self, other):
            empty_cell_equal_values = ['', 0, None, False]
            if other in empty_cell_equal_values:
                return True
            return False
