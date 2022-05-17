import sys

from excel2pycl import Cell, load_module


def load_module_with_cells(module_path, cells):
    return load_module(module_path).ExcelInPython(cells)


def main():
    module_path = sys.argv[1]

    cell = Cell(0, 2, 0)
    excel_class = load_module_with_cells(module_path, [Cell(0, 1, 20, 20000).to_dict()])
    print(excel_class.exec_function_in(cell.uid))


if __name__ == '__main__':
    main()
