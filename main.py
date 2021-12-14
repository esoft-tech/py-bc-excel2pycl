from src.context import Context
from src.excel import Excel, Cell
from src.translators import CellTranslator


def write_class_to_file(code) -> str:
    import sys
    filename = f'c_{hash(code) % ((sys.maxsize + 1) * 2)}'
    with open(f'temp/{filename}.py', 'w') as f:
        f.write(code)

    return filename


def exec_from_file_with_cells(filename, cells, exec_cell):
    from src.object_loader import load_object
    return load_object(f'temp.{filename}', 'ExcelInPython')(cells).exec_function_in(exec_cell)


if __name__ == '__main__':
    excel = Excel.parse('./test.xlsx')
    cell = Cell(0, 1, 0)
    context = Context()
    CellTranslator.translate(cell, excel, context)
    filename = write_class_to_file(context.build_class([Cell(0, 1, 10)], excel))
    print(exec_from_file_with_cells(filename, {Cell(0, 1, 10).uid: 125}, cell))

