from src.excel import Excel, Cell
from src.lexer import Lexer

if __name__ == '__main__':
    excel = Excel.parse('./test.xlsx')
    cell = Cell(0, 1, 0)
    data = excel.get_cell(cell)
    tokens = Lexer.parse(data)
    print(tokens)