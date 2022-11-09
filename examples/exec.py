import sys

from excel2pycl import Cell, Executor


def main():
    class_file = '1.py'

    print(Executor()
          .set_executed_class(class_file=class_file)
          .set_cells([Cell("Лист1", 3, 0, 48)])
          .get_cell(Cell("Лист1", 0, 0))
          .value)


if __name__ == '__main__':
    main()
