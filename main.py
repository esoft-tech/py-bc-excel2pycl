from src.context import Context
from src.excel import Excel, Cell
from src.translators import CellTranslator


if __name__ == '__main__':
    excel = Excel.parse('./test.xlsx')
    cell = Cell(0, 1, 0)
    context = Context()
    CellTranslator.translate(cell, excel, context)
    print(context.build_class([Cell(0, 1, 10)], excel))
