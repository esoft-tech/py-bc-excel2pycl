from abc import ABC
from typing import Dict


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

    @staticmethod
    def _binary_search(arr: list, x: any, reverse: bool = False):
        first = 0
        last = len(arr) - 1
        value = {
            'next_smallest': last if reverse else first,
            'next_largest': first if reverse else last,
            'exact': -1
        }

        while first <= last:

            mid = (last + first) // 2
            left = arr[mid] > x if reverse else arr[mid] < x
            right = arr[mid] < x if reverse else arr[mid] > x

            if left:
                if reverse:
                    value['next_largest'] = mid
                else:
                    value['next_smallest'] = mid

                first = mid + 1

            elif right:
                if reverse:
                    value['next_smallest'] = mid
                else:
                    value['next_largest'] = mid

                last = mid - 1

            else:
                value['exact'] = mid
                break

        if arr[value['next_smallest']] > x:
            value['next_smallest'] = -1

        if arr[value['next_largest']] < x:
            value['next_largest'] = -1

        return value

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
