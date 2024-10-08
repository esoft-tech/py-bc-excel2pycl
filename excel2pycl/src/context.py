from typing import Dict, List

from excel2pycl.src.cell import Cell


class Context:
    """
    Will be storing cell translation map.
    """

    def __init__(self):
        self._cell_translations: Dict[str, str] = {}
        self._sub_cell_translations: Dict[str, List] = {}
        self._titles: Dict[str, int] = {}
        self._sheets_size: List[Dict[str, int]] = []

    @property
    def __class_template(self) -> str:
        # TODO можно сделать кэш ячеек просчитанных
        return '''import datetime
from dateutil import parser as date_parser
from dateutil.relativedelta import relativedelta
from math import trunc, ceil, floor
from typing import Dict, List, Literal, Any, Callable
import calendar
import re
from itertools import zip_longest

class ExcelInPython:
    class ExcelInPythonException(Exception):
        pass
        
    def __init__(self, arguments: List = None):
        if arguments is None:
            arguments = []
        self._arguments: Dict[str, Any] = {{}}
        self.set_arguments(arguments)
        self._titles: Dict[str, int] = {titles}
        self._sheets_size: List[Dict[str, int]] = {sheets_size}

    def set_arguments(self, arguments: List):
        self._arguments = {{
            **self._arguments,
            **{{i['uid']: i['value'] for i in arguments}}
        }}

    def get_titles(self) -> Dict[str, int]:
        return self._titles
        
    def get_sheets_size(self) -> List[Dict[str, int]]:
        return self._sheets_size

    class EmptyCell(int):
        def __eq__(self, other: Any) -> bool:
            empty_cell_equal_values = ['', 0, None, False]
            if other in empty_cell_equal_values:
                return True
            
            return isinstance(other, self.__class__)
        
        def __lt__(self, other: Any) -> bool:
            if isinstance(other, (datetime.date, datetime.datetime)):
                # Ну вот так excel себя чувствует, пустая ячейка меньше любой даты
                return True
            
            if isinstance(other, (int, float)):
                return other > 0
            
            if isinstance(other, str):
                return other != ''
            
            if isinstance(other, list):
                return not other
            
            return False
        
        def __le__(self, other: Any) -> bool:
            return self.__eq__(other) or self.__lt__(other)
        
        def __gt__(self, other: Any) -> bool:
            return False
        
        def __ge__(self, other: Any) -> bool:
            return self.__eq__(other) or self.__gt__(other)
            
    def _parse_date_obj(self, date: str | datetime.datetime) -> datetime.datetime | None:
        if isinstance(date, datetime.datetime):
            return date

        try:
            return date_parser.parse(date)
        except (date_parser.ParserError, TypeError):
            return None

    def _by_operator(self, operator: str, left_operand: str | int | float | datetime.datetime, right_operand: str | int | float | datetime.datetime) -> bool:
        match operator:
            case '>=':
                return left_operand >= right_operand
            case '>':
                return left_operand > right_operand
            case '<=':
                return left_operand <= right_operand
            case '<':
                return left_operand < right_operand
            case '==':
                return left_operand == right_operand
            case '!=':
                return left_operand != right_operand
            case _:
                raise self.ExcelInPythonException('unknown operator ' + operator)

    def _compare(self, operator: str, left_operand: str | int | float | datetime.date | datetime.datetime,
                          right_operand: str | int | float | datetime.date | datetime.datetime) -> bool:
        try:
            return self._by_operator(operator, int(left_operand), int(right_operand))
        except (ValueError, TypeError):
            try:
                return self._by_operator(operator, float(left_operand), float(right_operand))
            except (ValueError, TypeError):
                try:
                    # Приводим date к datetime для удобного сравнения
                    if isinstance(left_operand, datetime.date) and not isinstance(left_operand, datetime.datetime):
                        left_operand = datetime.datetime(left_operand.year, left_operand.month, left_operand.day)
                    
                    if isinstance(right_operand, datetime.date) and not isinstance(right_operand, datetime.datetime):
                        right_operand = datetime.datetime(right_operand.year, right_operand.month, right_operand.day)
                    
                    return self._by_operator(operator, left_operand, right_operand)
                except (ValueError, TypeError):
                    return self._by_operator(operator, str(left_operand), str(right_operand))


    def _flatten_list(self, subject: List) -> List:
        result = []
        for i in subject:
            if isinstance(i, list):
                result = result + self._flatten_list(i)
            else:
                result.append(i)

        return result

    def _find_error_in_list(self, flatten_list: List):
        for err_value in filter(lambda cell: cell in ['#NUM!', '#DIV/0!',
                                                      '#N/A', '#NAME?', ' #NULL!',
                                                      '#REF!', '#VALUE!'], flatten_list):
            return err_value

    def _concat_arrays_values(self, list1: list, list2: list):
        return [[[str(x) + str(y)]] for x, y in zip_longest(list1, list2, fillvalue='')]

    def _normalize_float_number(self, number: float):
        return float(f'{{number:.15g}}')

    @staticmethod
    def _only_numeric_list(flatten_list: List, with_string_digits: bool = False):
        return [
            i 
            for i in flatten_list
            if type(i) in [float, int] or (with_string_digits and isinstance(i, str) and i.isdigit())
        ]

    @staticmethod
    def _only_bool_list(flatten_list: List):
        return [i for i in flatten_list if isinstance(i, bool)]
    
    @staticmethod
    def _only_datetime_list(flatten_list: List):
        return [i for i in flatten_list if isinstance(i, datetime.datetime)]
        
    @staticmethod
    def _regexp(pattern: str):
        pattern_flags = r'(?<![~])[?]+|[*]+'
        for item in re.finditer(pattern_flags, pattern):
            match item:
                case item if '?' in item.group():
                    pattern = pattern.replace(item.group(), '.' + '{{' + str(item.span()[1]-item.span()[0]) + '}}', 1)
                case item if '*' in item.group():
                    pattern = pattern.replace(item.group(), '.*', 1)
        pattern = re.sub(r'(?<=~)[?*]', r'\\\\\g<0>', pattern)
        pattern = re.sub(r'[\[\]]', r'\\\\\g<0>', pattern)
        return pattern

    @staticmethod
    def _binary_search(arr: List, lookup_value: any, reverse: bool = False):
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

    def _sum(self, flatten_list: List):
        return sum(self._only_numeric_list(flatten_list))

    def _average(self, flatten_list: List):
        return self._sum(flatten_list) / len(self._only_numeric_list(flatten_list))
    
    def _count(self, matrices: List[List], args: List, args_cells: List):
        flattened_matrices = self._flatten_list(matrices)
        return len(
            self._only_numeric_list(
                flattened_matrices + args_cells
            ) + self._only_bool_list(  # false и true учитываются
                args
            ) + self._only_numeric_list(
                args, with_string_digits=True
            ) + self._only_datetime_list(
                flattened_matrices + args_cells + args
            )
        )

    def _match(self, lookup_value, lookup_array: List, match_type: int = 0):
        lookup_value_type = int if isinstance(lookup_value, self.EmptyCell) else type(lookup_value)

        match match_type:
            case 0:
                for index, value in enumerate(lookup_array):
                    if isinstance(value[0], self.EmptyCell) or not isinstance(value[0], lookup_value_type):
                        continue
                    if value[0].lower() == lookup_value.lower() if isinstance(value[0], str) else value[0] == lookup_value:
                        return index + 1
                return '#N/A'
            case match_type if match_type > 0:
                last_valid_index = '#N/A'
                for index, value in enumerate(lookup_array):
                    if isinstance(value[0], self.EmptyCell) or not isinstance(value[0], lookup_value_type):
                        continue
                    if value[0].lower() <= lookup_value.lower() if isinstance(value[0], str) else value[0] <= lookup_value:
                        last_valid_index = index + 1
                    else:
                        return last_valid_index
            case match_type if match_type < 0:
                last_valid_index = '#N/A'
                for index, value in enumerate(lookup_array):
                    if isinstance(value[0], self.EmptyCell) or not isinstance(value[0], lookup_value_type):
                        continue
                    if value[0].lower() >= lookup_value.lower() if isinstance(value[0], str) else value[0] >= lookup_value:
                        last_valid_index = index + 1
                    else:
                        return last_valid_index

    def _xmatch(self, lookup_value, lookup_array: List, match_mode: int = 0, search_mode: int = 1):
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

    def _vlookup(self, lookup_value: str | int | float, table_array: List, col_index_num: int, range_lookup: bool | int = False):
        if not isinstance(range_lookup, (bool, int)):
            return '#ERROR!'

        lookup_value_type = int if isinstance(lookup_value, self.EmptyCell) else type(lookup_value)
        last_valid_value = '#N/A'
        
        def is_number(value):
            return isinstance(value, (float, int))

        for row in table_array:
            if isinstance(row[0], self.EmptyCell) or not isinstance(row[0], lookup_value_type):
                if not is_number(row[0]) or not is_number(lookup_value):
                    continue

            if range_lookup:
                if row[0] <= lookup_value:
                    last_valid_value = row[col_index_num - 1]
                else:
                    return last_valid_value
            elif row[0] == lookup_value:
                return row[col_index_num - 1]

        return last_valid_value

    def _sum_if(self, range_: List, criteria: Callable, sum_range: List = None):
        result = 0
        range_, sum_range = self._flatten_list(range_), self._flatten_list(sum_range)
        for i in range(len(range_)):
            if i < len(sum_range) and criteria(range_[i]):
                result += sum_range[i] or 0

        return result

    def _round(self, number: float, num_digits: int):
        return round(number, int(num_digits))

    def _roundup(self, number: float, num_digits: int):
        factor = 10 ** num_digits
        return ceil(number * factor) / factor

    def _rounddown(self, number: float, num_digits: int):
        factor = 10 ** num_digits
        return floor(number * factor) / factor

    def _date(self, year: int, month: int, day: int):
        if isinstance(year, str):
            try:
                year = int(year)
            except:
                return '#NUM!'
        
        if isinstance(month, str):
            try:
                month = int(month)
            except:
                return '#NUM!'
            
        if isinstance(day, str):
            try:
                day = int(day)
            except:
                return '#NUM!'
    
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
        
        
    def _left(self, text, num_chars):
        if num_chars is None:
            return text[0]
        if num_chars < 0:
            return '#ERROR!'
        if not text:
            return self.EmptyCell()
        if len(text) < num_chars:
            return text
        return text[0:num_chars]

    def _mid(self, text, start_num, num_chars):
        if start_num < 1:
            return '#NUM!'
        if num_chars < 0:
            return '#VALUE!'
        if start_num > len(text):
            return self.EmptyCell()
        
        return text[start_num - 1:start_num + num_chars - 1]
    
    @staticmethod
    def _address(row: int, col: int, *args) -> str:
        from string import ascii_uppercase

        def get_col():
            array = []

            def get_col_recursive(letters, _col):
                parent = _col // len(letters)
                child = _col % len(letters)

                if parent > len(letters):
                    array.append(get_col_recursive(letters, parent))
                return (letters[parent - 1] if parent < len(letters) and parent else '') + (
                    letters[child - 1] if child else '')

            array.append(get_col_recursive(ascii_uppercase, col))
            return ''.join(array)

        if not args:
            return '$' + get_col() + '$' + str(row)

        ref_type, *args = args
        col_value = ''
        match ref_type:
            case '1':
                col_value = '$' + get_col() + '$' + str(row)
            case '2':
                col_value = get_col() + '$' + str(row)
            case '3':
                col_value = '$' + get_col() + str(row)
            case '4':
                col_value = get_col() + str(row)

        if args:
            a1_type, *args = args
            if a1_type == 'False':
                col_value = 'R' + str(row) + 'C' + str(col)
                match ref_type:
                    case '2':
                        col_value = 'R' + str(row) + 'C' + '[' + str(col) + ']'
                    case '3':
                        col_value = 'R' + '[' + str(row) + ']' + 'C' + str(col)
                    case '4':
                        col_value = 'R' + '[' + str(row) + ']' + 'C' + '[' + str(col) + ']'

        if args:
            sheet_name, *args = args
            col_value = sheet_name + '!' + col_value

        return col_value
    
    def _right(self, text, num_chars):
        if num_chars is None:
            return text[len(text) - 1]
        if num_chars < 0:
            return '#ERROR!'
        if not text:
            return self.EmptyCell()
        if len(text) < num_chars:
            return text
        return text[len(text) - num_chars:]

    def _or(self, flatten_list: List):
        return any(flatten_list)

    def _and(self, flatten_list: List):
        return all(flatten_list)

    def _min(self, flatten_list: List):
        err_value = self._find_error_in_list(flatten_list)
        if err_value:
            return err_value

        return min(self._only_numeric_list(flatten_list))

    def _max(self, flatten_list: List):
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
        except:
            return when_error
    
    def _when_cell_is_empty_cast_to_zero(self, iterable: List):
        return [0 if isinstance(i, self.EmptyCell) else i for i in iterable]
        
    def _averageifs(self, average_range: List[List], *range_and_criteria):
        class Undefined:
            pass

        # Ячейки в диапазоне, содержащие значение TRUE, оцениваются как 1; ячейки в диапазоне,
        # содержащие значение FALSE, оцениваются как 0 (ноль).
        _when_bool_cast_to_int = lambda l: [int(i) if isinstance(i, bool) else i for i in l]

        average_range = self._flatten_list(average_range)
        # If average_range is a blank or text value, AVERAGEIFS returns the #DIV0! error value.
        if not average_range or isinstance(average_range, str) or isinstance(average_range, self.EmptyCell):
            return '#DIV0!'

        # Если ячейки в average_range не могут быть преобразованы в числа, AVERAGEIFS возвращает #DIV0! значение ошибки.
        try:
            [int(i) for i in average_range]
        except:
            return '#DIV0!'

        # List[List[criteria_range, criteria]]
        range_and_criteria_zip = []
        for i in range_and_criteria:
            if not range_and_criteria_zip or len(range_and_criteria_zip[-1]) == 2:
                i = self._flatten_list(i)
                if len(average_range) != len(i):
                    raise self.ExcelInPythonException('Invalid averageifs range size')
                range_and_criteria_zip.append([_when_bool_cast_to_int(self._when_cell_is_empty_cast_to_zero(i))])
            else:
                range_and_criteria_zip[-1].append(i)

        for [_range, criteria] in range_and_criteria_zip:
            for i in range(len(_range)):
                if not criteria(_range[i]):
                    average_range[i] = Undefined()


        average_range = _when_bool_cast_to_int([i for i in average_range if not isinstance(i, Undefined)])
        if not average_range or [True for i in average_range if isinstance(i, self.EmptyCell)]:
            return '#DIV/0!'

        return self._average(average_range)
    
    def _countifs(self, count_range: List[List], count_condition: Callable, *range_n_criteria):
        # Если ячейка в диапазоне критериев пуста, COUNTIFS обрабатывает ее как значение 0.

        count_range = self._flatten_list(count_range)

        range_and_criteria_zip = []
        for i in range_n_criteria:
            if not range_and_criteria_zip or len(range_and_criteria_zip[-1]) == 2:
                i = self._flatten_list(i)
                if len(count_range) != len(i):
                    raise self.ExcelInPythonException('Invalid countifs range size')
                range_and_criteria_zip.append([self._when_cell_is_empty_cast_to_zero(i)])
            else:
                range_and_criteria_zip[-1].append(i)

        for [_range, criteria] in range_and_criteria_zip:
            for i in range(len(_range)):
                if not criteria(_range[i]):
                    count_range[i] = None
        count_range = [i if count_condition(i) else None for i in count_range]
        return len(list(filter(None, count_range)))
        
    def _network_days(self, date_start: datetime.datetime, date_end: datetime.datetime,
                      holidays: List[List[datetime.datetime]] | None = None):
        # Большая загадка как вычисляется значение если на входе не даты - поэтому я решила просто кидать '#VALUE!'
        if not isinstance(date_start, datetime.datetime) or not isinstance(date_end, datetime.datetime):
            return '#VALUE!'

        work_days_count = 0
        if date_start.date() <= date_end.date():
            start = date_start.date()
            end = date_end.date()
            multiple = 1
        else:
            start = date_end.date()
            end = date_start.date()
            multiple = -1

        additional_days = []
        if holidays:
            for row in holidays:
                additional_days_in_row = [day.date() for day in row if isinstance(day, datetime.datetime)] \\
                    if row is not None else []
                additional_days += additional_days_in_row

        while start <= end:
            if start.weekday() not in [5, 6] and start not in additional_days:
                work_days_count += 1
            start = start + datetime.timedelta(days=1)

        return work_days_count * multiple
        
    def _sumifs(self, sum_range: List[List], *range_and_criteria):
        # Ячейки в диапазоне, содержащие значение TRUE, оцениваются как 1; ячейки в диапазоне,
        # содержащие значение FALSE, оцениваются как 0 (ноль).
        _when_bool_cast_to_int = lambda l: [int(i) if isinstance(i, bool) else i for i in l]

        sum_range = self._flatten_list(sum_range)

        range_and_criteria_zip = []
        for i in range_and_criteria:
            if not range_and_criteria_zip or len(range_and_criteria_zip[-1]) == 2:
                i = self._flatten_list(i)
                if len(sum_range) != len(i):
                    raise self.ExcelInPythonException('Invalid sumifs range size')
                range_and_criteria_zip.append([_when_bool_cast_to_int(self._when_cell_is_empty_cast_to_zero(i))])
            else:
                range_and_criteria_zip[-1].append(i)

        for [_range, criteria] in range_and_criteria_zip:
            for i in range(len(_range)):
                if not criteria(_range[i]):
                    sum_range[i] = None

        sum_range = _when_bool_cast_to_int([i for i in sum_range if i is not None])

        return self._sum(sum_range)

    def _index(self, matrix_list: tuple | list, row_number: int, column_number: int | None, area_number: int):
        if area_number > len(matrix_list):
            return '#REF!'

        # Если пришел кортеж, значит имеем дело с несколькими диапазонами, берем заданный в area_number, по умолчанию 1
        array = matrix_list[area_number - 1] if isinstance(matrix_list, tuple) else matrix_list
        
        # Если диапазон - строка и указан только номер строки, считаем его номером столбца
        if len(array) == 1 and column_number is None:
            column_number = row_number
            row_number = None

        try:
            # Если не указаны номер столбца/строки, берем значения из всех столбцов/строк
            row = [array[row_number - 1]] if row_number else array
            value = [col[column_number - 1] if column_number else col for col in row]
        except IndexError:
            return '#REF!'

        # Для диапазона типа столбец значения будут в конструкции [[x], [y], [z]], приводим к аналогу строки - [x, y, z]
        if isinstance(value[0], list) and len(value[0]) == 1:
            value = [row[0] for row in value]

        return value[0] if len(value) == 1 else value

    def _cell_preprocessor(self, cell_uid: str):
        # Ищем метод расчета значения ячейки среди методов и аттрибутов экземпляра и класса
        method = self.__dict__.get(cell_uid, self.__class__.__dict__.get(cell_uid))
        # Ищем значение значение ячейки среди установленных в ручную через set_cells, если не находим, считаем результат
        # с помощью найденного выше метода, если же не найден и метод, возвращаем "пустую ячейку"
        return self._arguments.get(cell_uid, method(self) if method else self.EmptyCell())

    def exec_function_in(self, cell_uid: str):
        return self._cell_preprocessor(cell_uid)
    
    @staticmethod
    def _today() -> datetime.date:
        return datetime.datetime.combine(datetime.date.today(), datetime.time(0, 0))

        
    def _count_blank(self, flatten_list: List):
        err_value = self._find_error_in_list(flatten_list)
        if err_value:
            return err_value
        empty = [elem for elem in flatten_list if elem is None or elem == ""]
        return len(empty)

    def _ifs(self, flatten_list: List):
        err_value = self._find_error_in_list(flatten_list)
        if err_value:
            return err_value

        index = 0
        while index < len(flatten_list):
            if flatten_list[index]:
                return flatten_list[index + 1]
            index += 2

        return '#N/A'


    def _search(self, find_text: str, within_text: str, start_num: int | None):
        start_num = start_num if start_num else 1
        if start_num and (start_num > len(within_text) or start_num <= 0):
            return '#VALUE!'

        pattern = r'([^~][?*]|^[?*])'
        if len(re.findall(pattern, find_text)) == 0:
            find_text = find_text.replace('~?', '?') \
                .replace('~*', '*')

            result = within_text.find(find_text, start_num - 1) + 1
            return result if result else '#VALUE!'

        find_text = find_text \
            .replace('?', '(.)') \
            .replace('*', '(.*)') \
            .replace('~(.*)', r'\*') \
            .replace('~(.)', r'\?')

        result = re.finditer(find_text, within_text, re.I)

        if result is None:
            return '#VALUE!'

        find_elem = None
        for i in result:
            if i.span(0)[0] + 1 < start_num:
                continue
            find_elem = i
            break
        # исключаем поиск по regex вроде \d
        if find_elem:
            sequences = find_elem.groups(0)
            found_text = find_elem.group(0)
            find_text = find_text.replace('(.*)', '(.)') \
                                 .replace(r'\?', '?') \
                                 .replace(r'\.', '.')
            for sequence in sequences:
                find_text = find_text.replace('(.)', sequence, 1)
    
            if found_text.lower() != find_text.lower():
                return '#VALUE!'
            
        return find_elem.span(0)[0] + 1 if find_elem else '#VALUE!'

    def _excel_value_to_string(self, value: Any):
        if isinstance(value, (datetime.datetime)):
            base_date = datetime.datetime(1899, 12, 30)
            return str((value - base_date).days)

        return str(value)

    def _parse_date_formats(self, date: str, format: str):
        try:
            return datetime.datetime.strptime(date, format)
        except ValueError:
            return datetime.datetime.strptime(date, f'{{format}} 00:00:00')

    def _value(self, text: str):
        # Удаляем пробелы в начале и конце строки
        text = text.strip()

        # Попытка преобразовать строку в целое число
        try:
            return int(text)
        except ValueError:
            pass

        # Попытка преобразовать строку в число с плавающей точкой
        try:
            # Заменяем запятую на точку, если используется десятичная запятая
            text = text.replace(",", ".")
            return float(text)
        except ValueError:
            pass

        # Попытка преобразовать строку в дату в разных форматах
        date_formats = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"]
        base_date = datetime.datetime(1899, 12, 30)  # Базовая дата для Excel (1 января 1900 года = 1 день)
        for fmt in date_formats:
            try:
                date = self._parse_date_formats(text, fmt)
                # Возвращаем количество дней, прошедших с 1 января 1900 года
                return (date - base_date).days
            except ValueError:
                continue

        time_formats = ["%H:%M:%S", "%H:%M"]
        for fmt in time_formats:
            try:
                time_value = datetime.datetime.strptime(text, fmt).time()
                # Преобразуем время в долю дня
                return (time_value.hour + time_value.minute / 60 + time_value.second / 3600) / 24
            except ValueError:
                continue

        # Попытка преобразовать строку с процентами
        if text.endswith("%"):
            try:
                return float(text[:-1].replace(",", ".")) / 100
            except ValueError:
                pass

        # Удаление пробелов и разделителей тысяч
        # Пример: "1 234,56" -> "1234.56"
        clean_text = text.replace(" ", "").replace(",", ".")
        try:
            return float(clean_text)
        except ValueError:
            pass

        return '#VALUE!'

{functions}
'''

    @property
    def __function_template(self) -> str:
        return '''    def {name}(self):
        return {code}'''

    def __build_function(self, name: str, code: str) -> str:
        return self.__function_template.format(name=name, code=code)

    def __build_functions(self, functions: Dict) -> str:
        return '\n\n'.join([self.__build_function(name, code) for name, code in functions.items()])

    def __build_class(self, functions: Dict) -> str:
        return self.__class_template.format(functions=self.__build_functions(functions), titles=self._titles,
                                            sheets_size=self._sheets_size)

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

    def _get_divided_sub_cell_translations(self) -> Dict:
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
