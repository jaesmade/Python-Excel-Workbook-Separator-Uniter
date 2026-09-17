import os
from openpyxl import Workbook
from excel_splitter.core import split_workbooks

files = [
    os.path.join(os.getcwd(), 'dup_a.xlsx'),
    os.path.join(os.getcwd(), 'dup_b.xlsx'),
]
out = os.path.join(os.getcwd(), 'dup_output.xlsx')

for path, dept in [(files[0], 'ACCOUNTING OFFICE'), (files[1], 'ACCOUNTING OFFICE')]:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.append(['Department', 'Name', 'ID Number'])
    ws.append([dept, 'Employee A', 'PERMANENT'])
    ws.append([dept, 'Employee B', 'UNKNOWN'])
    wb.save(path)

result = split_workbooks(files, 'Sheet1', 'Department', 'ID Number', out, combine_duplicates=True)
print(result)
