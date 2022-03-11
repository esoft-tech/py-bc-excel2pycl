from excel2pycl import Context, Excel, Cell, ExcelSafeException, CellTranslator, load_object


def write_class_to_file(code) -> str:
    import sys
    filename = f'c_{hash(code) % ((sys.maxsize + 1) * 2)}'
    with open(f'temp/{filename}.py', 'w') as f:
        f.write(code)

    return filename


def exec_from_file_with_cells(filename, cells, exec_cell):
    return load_object(f'temp.{filename}', 'ExcelInPython')(cells).exec_function_in(exec_cell)


if __name__ == '__main__':
    excel = Excel.parse('./test.xlsx')
    try:
        excel.is_safe()
    except ExcelSafeException as ese:
        print(ese)
    cell = Cell(0, 2, 0)
    context = Context()
    CellTranslator.translate(cell, excel, context)
    filename = write_class_to_file(context.build_class([], excel))
    print(exec_from_file_with_cells(filename, {}, cell))
    print(filename)
