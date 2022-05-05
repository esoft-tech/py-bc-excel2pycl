from excel2pycl.src.cell import Cell
from excel2pycl.src.excel import Excel


class Context:
    def __init__(self):
        self._cell_translations = {}
        self._sub_cell_translations = {}

    @property
    def __class_template(self) -> str:
        # TODO можно сделать кэш ячеек просчитанных
        return '''class ExcelInPython:
    def __init__(self, arguments):
        for key, value in arguments.items():
            setattr(self, key, lambda _=None: value)
        
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
        return self._sum(flatten_list)/len(self._only_numeric_list(flatten_list))
        
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

    def exec_function_in(self, cell):
        return self.__class__.__dict__[cell.uid](self)

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
        return self.__class_template.format(functions=self.__build_functions(functions))

    @staticmethod
    def _get_cell_function_name(cell: Cell) -> str:
        # TODO Придумать как сделать это частью cell
        return cell.uid

    @classmethod
    def _get_sub_cell_function_name(cls, cell: Cell = None, sub_number: int = None, cell_prefix: str = None) -> str:
        return f'{cell_prefix or cls._get_cell_function_name(cell)}_{sub_number}'

    def get_cell(self, cell: Cell) -> str or None:
        return f'self.{self._get_cell_function_name(cell)}()' if cell.uid in self._cell_translations else None

    def set_cell(self, cell: Cell, code: str) -> str:
        self._cell_translations[self._get_cell_function_name(cell)] = code
        return self.get_cell(cell)

    def set_sub_cell(self, cell: Cell, code: str) -> str:
        # TODO check if sub expression exists
        cell_function_name = self._get_cell_function_name(cell)
        if not self._sub_cell_translations.get(cell_function_name):
            self._sub_cell_translations[self._get_cell_function_name(cell)] = []
        self._sub_cell_translations[cell_function_name].append(code)

        return f'self.{self._get_sub_cell_function_name(cell=cell, sub_number=len(self._sub_cell_translations[cell_function_name]) - 1)}()'

    def _get_divided_sub_cell_translations(self) -> dict:
        result = {}
        for cell_prefix, sub_cell_expressions in self._sub_cell_translations.items():
            for sub_cell_expression, sub_cell_index in zip(sub_cell_expressions, range(len(sub_cell_expressions))):
                result[self._get_sub_cell_function_name(cell_prefix=cell_prefix, sub_number=sub_cell_index)] = sub_cell_expression

        return result

    def build_class(self, argument_cells: list, excel: Excel) -> str:
        for cell in argument_cells:
            excel.handle_cell(cell)
            self.set_cell(cell, f'self._arguments[\'{cell.uid}\']')

        summary_functions = self._cell_translations.copy()
        summary_functions.update(self._get_divided_sub_cell_translations())

        return self.__build_class(summary_functions)

