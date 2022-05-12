import sys

from excel2pycl import Context, Excel, Cell, ExcelSafeException, CellTranslator


def write_class_to_file(code: str) -> str:
    """
    Write class function example
    """
    filename = f'c_{hash(code) % ((sys.maxsize + 1) * 2)}'
    with open(f'temp/{filename}.py', 'w') as f:
        f.write(code)

    return filename


def main():
    # Parsing Excel file content
    excel = Excel.parse('./test.xlsx')
    try:
        # Checking the file content is safety
        excel.is_safe()
    except ExcelSafeException as ese:
        print(ese)

    # Initializing entrypoint cell for Excel file
    cell = Cell(0, 2, 0)
    # Initializing Content instance, where will be storing cell translation map
    context = Context()

    # Translating all cells needed for executing entrypoint cell to python class
    CellTranslator.translate(cell, excel, context)
    # Writing translated class to the python file
    print(write_class_to_file(context.build_class()))


if __name__ == '__main__':
    main()
