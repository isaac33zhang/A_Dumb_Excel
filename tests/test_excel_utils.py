# -*- coding: utf-8 -*-
import os
from lib.excel_utils import ExcelObject


FILE_PATH = "/Users/fanzhaf/Documents/Personal/pangpang/actual_departments.xlsx"


def test_get_sheet_objects():
    eo = ExcelObject(FILE_PATH)
    sheets = eo.get_sheet_objects()
    assert len(sheets) == 1
    assert sheets[0].title.encode('utf-8') == "局"


def test_get_cell_value():
    eo = ExcelObject(FILE_PATH)
    cell_value = eo.get_cell(2, 2)
    assert cell_value.encode('utf-8') == "王静"


def test_get_row_count():
    eo = ExcelObject(FILE_PATH)
    row_count = eo.get_row_count()
    assert row_count == 189


def test_get_column_count():
    eo = ExcelObject(FILE_PATH)
    column_count = eo.get_column_count()
    assert column_count == 7


def test_get_row_list():
    eo = ExcelObject(FILE_PATH)
    row_list = eo.get_row_list(2)
    assert row_list[1].value.encode('utf-8') == "王静"
    assert row_list[6].value.encode('utf-8') == "安定门所"


def test_create_file():
    file_path = "/Users/fanzhaf/Desktop/测试.xlsx"
    ExcelObject.create_file(file_path)
    wb = ExcelObject(file_path)
    assert wb.workbook
    os.remove(file_path)


