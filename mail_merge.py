from segmentAdder import segmentAdder
from openpyxl import load_workbook
import re

def process_merge(data, flightSegments):
    """
    :param data: a dict of values, keys being the <<Values>> in lowercase form
    :param flightSegmentData: array of dicts corresponding to flight legs
    :return: True if all went well
    """
    pattern = re.compile(r'(<<)(\w+)(>>)')
    wb = load_workbook('/Users/kyleblazepetan/Desktop/mailMergeTemplate.xlsx')
    ws = wb.active
    segments = len(flightSegments)

    def lookUpFunction(matchobj):
        resp = data[str(matchobj.group(2)).lower()]
        if resp:
            return resp
        return 'key not found'
    
    segmentAdder(segments, ws)

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.value = re.sub(pattern, lookUpFunction, cell.value.lower())

    wb.save('/Users/kyleblazepetan/Desktop/mergeTesting.xlsx')
