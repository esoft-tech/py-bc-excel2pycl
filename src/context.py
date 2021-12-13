from src.cell import Cell
from src.excel import Excel


class Context:
    def __init__(self):
        self._cell_translations = {}
        self._sub_cell_translations = {}

    @property
    def __class_template(self) -> str:
        # TODO можно сделать кэш ячеек просчитанных
        return '''class ExcelInPython:
    def __init__(self, arguments):
        self._arguments = arguments

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
    def _get_sub_cell_function_name(cls, cell: Cell, sub_number: int) -> str:
        return f'{cls._get_cell_function_name(cell)}_{sub_number}'

    def get_cell(self, cell: Cell) -> str or None:
        return f'self.{self._get_cell_function_name(cell)}()' if cell.uid in self._cell_translations else None

    def set_cell(self, cell: Cell, code: str) -> str:
        self._cell_translations[self._get_cell_function_name(cell)] = code
        return self.get_cell(cell)

    def set_sub_cell(self, cell: Cell, code: str) -> str:
        cell_function_name = self._get_cell_function_name(cell)
        if not self._sub_cell_translations.get(cell_function_name):
            self._sub_cell_translations[self._get_cell_function_name(cell)] = []
        self._sub_cell_translations[cell_function_name].append(code)

        return f'self.{self._get_sub_cell_function_name(cell, len(self._sub_cell_translations[cell_function_name]) - 1)}()'

    def build_class(self, argument_cells: list, excel: Excel) -> str:
        for cell in argument_cells:
            excel.handle_cell(cell)
            self.set_cell(cell, f'self._arguments[\'{cell.uid}\']')

        summary_functions = self._cell_translations.copy()
        summary_functions.update(self._sub_cell_translations)

        return self.__build_class(summary_functions)

