from typing import Dict

from openpyxl.utils import column_index_from_string

from excel2pycl.src.cell import Cell


def handle_cell(cell: Cell, titles: Dict[str, int]):
    if cell.has_handled_identifiers():
        return

    if type(cell.title) is str:
        cell.title = titles[cell.title]

    if type(cell.column) is str:
        cell.column = column_index_from_string(cell.column) - 1

    if type(cell.row) is str:
        if cell.row:
            cell.row = int(cell.row) - 1
        else:
            cell.row = None

    cell._handled_identifiers = True
