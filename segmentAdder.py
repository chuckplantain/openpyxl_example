from openpyxl import load_workbook
from itertools import *
import re

pattern = re.compile(r'(Flight Segment )(\d*)')

def segment_repl(matchobj):
    base = str(matchobj.group(0))
    digits = str(matchobj.group(1))
    return base + ' ' + '34343'

wb = load_workbook('/Users/kyleblazepetan/Desktop/mailMergeTemplate.xlsx')
ws = wb.active
segments = 55 + ( 23 * 3 )
for index in xrange(55, segments, 23):
    pertinent_rows = ws['A32':'L54']
    blank_rows = ws.iter_rows(row_offset=index)
    for row, bl_row in zip(pertinent_rows, blank_rows):
        for cell, bl_cell in zip(row, bl_row):
            desired_style = cell.style
            if cell.value is not None:
                bl_cell.value = re.sub(pattern, segment_repl, cell.value)
                bl_cell.style = desired_style

wb.save('/Users/kyleblazepetan/Desktop/segements.xlsx')
