import unittest
import uuid
import os
from excel2pycl import Parser, Executor, Cell
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo



def create_test_table(file_name):
    wb = Workbook()
    ws = wb.active

    data = [
        [93, '=IFS(A1>89;"A";A1>79;"B";A1>69;"C";A1>59;"D")'],
        [89, '=IFS(A2>89;"A";A2>79;"B";A2>69;"C";A2>59;"D")'],
        [71, '=IFS(A3>89;"A";A3>79;"B";A3>69;"C";A3>59;"D")'],
        [60, '=IFS(A4>89;"A";A4>79;"B";A4>69;"C";A4>59;"D")'],
        [58, '=IFS(A5>89;"A";A5>79;"B";A5>69;"C";A5>59;"D")'],
    ]

    for row in data:
        ws.append(row)

    tab = Table(displayName='base', ref='A1:B5')

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws.add_table(tab)

    ws_2 = wb.create_sheet('example')

    data = [
        ['', 'ПРОВЕРКА РИЕЛТОРА НА МОМЕНТ РАСЧЁТА:', 'Значение:', 'На что влияет:'],
        ['1.', 'Риелтор на стипендии?', 'нет', 'Фиксированные проценты за первые 5 сделок.'],
        ['2.', 'Риелтор является менеджером?', 'да', 'Процент менеджерам зафиксирован.'],
        ['3.', 'У риелтора фиксированный процент?', 'нет', 'Процент зафиксирован в каталоге.'],
        ['4.', 'Риелтор группы 0?', 'нет', 'При нахождении в группе 0 понижающий KPI не применяется.'],
        ['5.', 'Есть минимальный процент на рассчитываемый месяц?', 5 ,'Применяется если рассчитанный процент ниже указанного'],
        ['', 'РАССЧИТАННЫЕ ПОКАЗАТЕЛИ:'],
        ['1.', 'Кол-во сделок (всего)', 1],
        ['2.', 'Сумма валовой прибыли (за последние 3 мес.)', 0],
        ['3.', 'Процент выполнения KPI', 0],
        ['4.', 'Кол-во дней отпуска (для корректировки границ ВВ)', 0],
        ['', ''],
        ['', 'КОЭФФИЦИЕНТЫ ДЛЯ РАСЧЁТА ИТОГОВОГО РИЕЛТОРСКОГО ПРОЦЕНТА:'],
        ['1.', 'Базовый процент', 'фикс %'],
        ['2.', 'Понижающий KPI', 'фикс %'],
        ['3.', 'Процент по валовой прибыли', 'фикс %'],
        ['4.', 'Добавочный процент от 100 сделок', 'фикс %'],
        ['5.', 'Итоговый риелторский процент ', 50],
        ['', 'Применён минимальный?',
         '=IFS(C6=6;"нет";C18>C6;"нет";SUM(C$14:C$17)=C6;"нет";SUM(C$14:C$17)<>C6;"да")']
    ]

    for row in data:
        ws_2.append(row)

    tab = Table(displayName='base1', ref='A1:C30')

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style

    ws_2.add_table(tab)

    wb.save(file_name)

    wb.close()


class TestIfsToken(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.translation_file_path = f'test/{uuid.uuid4()}.py'
        create_test_table('test/ifs.xlsx')
        cls.parser = Parser() \
            .set_excel_file_path('test/ifs.xlsx') \
            .enable_safety_check() \
            .write_translation(cls.translation_file_path)

    @classmethod
    def tearDownClass(cls) -> None:
        # после выполнения всех тестов удаляем файлики
        os.remove(cls.translation_file_path)
        os.remove('test/ifs.xlsx')

    def test_true_in_first(self):
        excepted_cell_value = 'A'
        cell_value = Executor() \
            .set_executed_class(class_file=self.translation_file_path) \
            .get_cell(Cell(0, 1, 0))

        self.assertEqual(cell_value.value, excepted_cell_value, msg='test_true_in_first')

    # def test_true_in_last(self):
    #     excepted_cell_value = 'D'
    #     cell_value = Executor() \
    #         .set_executed_class(class_file=self.translation_file_path) \
    #         .get_cell(Cell(0, 1, 3))
    #
    #     self.assertEqual(cell_value.value, excepted_cell_value, msg='test sell concat')
    #
    # def test_true_in_middle(self):
    #     excepted_cell_value = 'C'
    #     cell_value = Executor() \
    #         .set_executed_class(class_file=self.translation_file_path) \
    #         .get_cell(Cell(0, 1, 2))
    #
    #     self.assertEqual(cell_value.value, excepted_cell_value)
    #
    # def test_not_true(self):
    #     excepted_cell_value = '#N/A'
    #     cell_value = Executor() \
    #         .set_executed_class(class_file=self.translation_file_path) \
    #         .get_cell(Cell(0, 1, 4))
    #
    #     self.assertEqual(cell_value.value, excepted_cell_value)

    # def test_some(self):
    #     excepted_cell_value = 'нет'
    #     cell_value = Executor() \
    #         .set_executed_class(class_file=self.translation_file_path) \
    #         .get_cell(Cell(1, 2, 18))
    #
    #     print(cell_value.value)
    #
    #    self.assertEqual(cell_value.value, excepted_cell_value)
