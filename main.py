from openpyxl.reader.excel import load_workbook

import toexcel
import sys
import os
import pandas as pd

file_name = 'output.xlsx'
# 检查文件是否存在
if os.path.exists(file_name):
    # 删除文件
    print('发现目标文件已存在，执行删除')
    os.remove(file_name)
file_name = 'output.xlsx'
df = pd.DataFrame()
df.to_excel(file_name, index=False)
print('空表格创建完毕')

lines = []
print("请输入SQL语句，按Ctrl+D(windows环境下Ctrl+Z结束)表示输入结束：")
for line in sys.stdin:
    nowl = line.rstrip().lower()
    if len(nowl) == 0:
        toexcel.insertexcel(lines)
        lines = []
    lines.append(line.rstrip().lower())
if len(lines) > 0:
    toexcel.insertexcel(lines)

wb = load_workbook(file_name)
del wb['Sheet1']
wb.save(file_name)

print("创建完成")
