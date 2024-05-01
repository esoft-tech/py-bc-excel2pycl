from typing import Dict

from openpyxl.utils import column_index_from_string

from excel2pycl.src.cell import Cell


def handle_cell(cell: Cell, titles: Dict[str, int]) -> None:
    if cell.has_handled_identifiers():
        return

    if isinstance(cell.title, str):
        cell.title = titles[cell.title]

    if isinstance(cell.column, str):
        cell.column = column_index_from_string(cell.column) - 1

    if isinstance(cell.row, str):
        if cell.row:
            cell.row = int(cell.row) - 1
        else:
            cell.row = None

    cell._handled_identifiers = True
