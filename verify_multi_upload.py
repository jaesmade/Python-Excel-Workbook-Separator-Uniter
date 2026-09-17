import os

from openpyxl import Workbook

from excel_splitter.core import split_workbooks

folder = os.getcwd()
file1 = os.path.join(folder, 'multi_a.xlsx')
file2 = os.path.join(folder, 'multi_b.xlsx')
out = os.path.join(folder, 'multi_output.xlsx')

for path, department in [(file1, 'ACCOUNTING OFFICE'), (file2, 'HR OFFICE')]:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.append(['Department', 'Name', 'ID Number'])
    ws.append([department, 'Employee A', 'PERMANENT'])
    ws.append([department, 'Employee B', 'UNKNOWN'])
    wb.save(path)

result = split_workbooks([file1, file2], 'Sheet1', 'Department', 'ID Number', out)
print(result)
print('out exists', os.path.exists(out))
