from abc import ABC
import datetime
import calendar
from dateutil.relativedelta import relativedelta
from typing import Dict, Literal
from math import trunc

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

    @staticmethod
    def _binary_search(arr: list, lookup_value: any, reverse: bool = False):
        first = 0
        last = len(arr) - 1
        next_smallest = last if reverse else first
        next_largest = first if reverse else last
        exact = -1

        while first <= last:

            mid = (last + first) // 2
            left = arr[mid][0] > lookup_value if reverse else arr[mid][0] < lookup_value
            right = arr[mid][0] < lookup_value if reverse else arr[mid][0] > lookup_value

            if left:
                if reverse:
                    next_largest = mid
                else:
                    next_smallest = mid

                first = mid + 1

            elif right:
                if reverse:
                    next_smallest = mid
                else:
                    next_largest = mid

                last = mid - 1

            else:
                exact = mid
                next_smallest = mid
                next_largest = mid
                break

        if arr[next_smallest][0] > lookup_value:
            next_smallest = -1

        if arr[next_largest][0] < lookup_value:
            next_largest = -1

        return (exact, next_smallest, next_largest)

    def _sum(self, flatten_list: list):
        return sum(self._only_numeric_list(flatten_list))

    def _average(self, flatten_list: list):
        return self._sum(flatten_list) / len(self._only_numeric_list(flatten_list))

    def _match(self, lookup_value, lookup_array: list, match_type: int = 0):
        lookup_value_type = int if isinstance(lookup_value, self.EmptyCell) else type(lookup_value)

        match match_type:
            case 0:
                for index, value in enumerate(lookup_array):
                    if isinstance(value[0], self.EmptyCell) or not isinstance(value[0], lookup_value_type):
                        continue
                    if value[0] == lookup_value:
                        return index + 1
                return '#N/A'
            case match_type if match_type > 0:
                last_valid_index = '#N/A'
                for index, value in enumerate(lookup_array):
                    if isinstance(value[0], self.EmptyCell) or not isinstance(value[0], lookup_value_type):
                        continue
                    if value[0] <= lookup_value:
                        last_valid_index = index + 1
                    else:
                        return last_valid_index
            case match_type if match_type < 0:
                last_valid_index = '#N/A'
                for index, value in enumerate(lookup_array):
                    if isinstance(value[0], self.EmptyCell) or not isinstance(value[0], lookup_value_type):
                        continue
                    if value[0] >= lookup_value:
                        last_valid_index = index + 1
                    else:
                        return last_valid_index

    def _xmatch(self, lookup_value, lookup_array: list, match_mode: int = 0, search_mode: int = 1):
        # TODO wildcard match
        output_value = 0
        match match_mode:
            case -1:
                output_value = 1
            case 1:
                output_value = 2

        match search_mode:
            case 1:
                return self._match(lookup_value, lookup_array, match_mode)
            case -1:
                return self._match(lookup_value, lookup_array[::-1], match_mode)
            case 2:
                index = self._binary_search(lookup_array, lookup_value)[output_value]
                return index + 1 if index != -1 else '#N/A'
            case -2:
                index = self._binary_search(lookup_array, lookup_value, reverse=True)[output_value]
                return index + 1 if index != -1 else '#N/A'
            case _:
                return '#ERROR!'

    def _vlookup(self, lookup_value, table_array: list, col_index_num: int, range_lookup: bool = False):
        if not isinstance(range_lookup, (bool, int)):
            return '#ERROR!'

        lookup_value_type = int if isinstance(lookup_value, self.EmptyCell) else type(lookup_value)

        if range_lookup:
            last_valid_value = '#N/A'
            for row in table_array:
                if isinstance(row[0], self.EmptyCell) or not isinstance(row[0], lookup_value_type):
                    continue
                if row[0] <= lookup_value:
                    last_valid_value = row[col_index_num - 1]
                else:
                    return last_valid_value
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

    def _edate(self, start_date: datetime, months: float):
        if not isinstance(start_date, datetime):
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

    class EmptyCell(int):
        def __eq__(self, other):
            empty_cell_equal_values = ['', 0, None, False]
            if other in empty_cell_equal_values:
                return True
            return False
