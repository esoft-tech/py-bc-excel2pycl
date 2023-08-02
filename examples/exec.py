import sys

from excel2pycl import Cell, Executor


def main():
    class_file = sys.argv[1]

    print(Executor()
          .set_executed_class(class_file=class_file)
          .set_cells([Cell(0, 0, 20, 20000)])
          .get_cell(Cell(0, 2, 0))
          .value)


if __name__ == '__main__':
    main()
