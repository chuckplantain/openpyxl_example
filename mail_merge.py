from openpyxl import load_workbook
from manifest import process_manifest
from insert_into_xlxs_from_db import write_xlsx_file
import re

data = {}
data['foo'] = "Kyle Petan"
data['bar'] = "Splendid and Great"

pattern = re.compile(r'(<<)(\w+)(>>)')

def lookUpFunction(matchobj):
    resp = data_dict[str(matchobj.group(2)).lower()]
    if resp:
        return resp
    return 'key not found'

def process_merge(data, template_file, new_file):
    """

    :param data: a dict of values, keys being the <<Values>> in lowercase form
    :param template_file: the file with the template (.xlsx)
    :param new_file: A place to save the newFile... ends in .xlsx and destructively saves
    :return: True if all went well
    """
    wb = load_workbook(template_file)
    ws = wb.active


    def lookUpFunction(matchobj):
        resp = data[str(matchobj.group(2)).lower()]
        if resp:
            return resp
        return 'key not found'


    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.value = re.sub(pattern, lookUpFunction, cell.value.lower())
    wb.save(new_file)

    return True

process_merge(data, '/Users/kyleblazepetan/Plumb/xlsxFiles/testFile.xlsx', '/Users/kyleblazepetan/tmp/shit.xlsx')
