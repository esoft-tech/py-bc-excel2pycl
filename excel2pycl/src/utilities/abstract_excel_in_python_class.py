from abc import ABC
import datetime
import calendar
from typing import Dict, Literal
from dateutil.relativedelta import relativedelta


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

    def _date(self, year: int, month: int, day: int):
        match year:
            case year if 0 <= year <= 1899:
                year += 1900
            case year if year < 0 or year > 9999:
                return '#NUM!'

        result_date = datetime(year, 1, 1)

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

    def _day(self, date: datetime):
        return date.day

    def _month(self, date: datetime):
        return date.month

    def _year(self, date: datetime):
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

    class EmptyCell(int):
        def __eq__(self, other):
            empty_cell_equal_values = ['', 0, None, False]
            if other in empty_cell_equal_values:
                return True
            return False
