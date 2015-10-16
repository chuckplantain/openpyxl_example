from openpyxl import load_workbook
import re

def process_merge(data, flightSegmentData):
    """

    :param data: a dict of values, keys being the <<Values>> in lowercase form
    :param template_file: the file with the template (.xlsx)
    :param new_file: A place to save the newFile... ends in .xlsx and destructively saves
    :return: True if all went well
    """
    pattern = re.compile(r'(<<)(\w+)(>>)')
    wb = load_workbook('/Users/kyleblazepetan/Desktop/mailMergeTemplate.xlsx')
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

    wb.save('/Users/kyleblazepetan/Desktop/maergeTesting.xlsx')
    return True
