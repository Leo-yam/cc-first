# -*- coding: utf-8 -*-
"""Generate test data using Unicode escapes (pure ASCII-safe)."""
from openpyxl import Workbook
import codecs

# All Chinese text via Unicode escape sequences
names_data = [
    '张三',       # 张三
    '李四',       # 李四
    '王小明', # 王小明
    '欧阳锋', # 欧阳锋
    '赵六六', # 赵六六
    '孙七',       # 孙七
    '周八',       # 周八
    '钱九',       # 钱九
    '刘磊',       # 刘磊
    '陈静',       # 陈静
]

# --- Excel ---
wb = Workbook()
ws = wb.active
ws.title = '党员名单'  # 党员名单

headers = [
    '序号',          # 序号
    '党员姓名',  # 党员姓名
    '支部',          # 支部
    '备注',          # 备注
]
for j, h in enumerate(headers):
    ws.cell(row=1, column=j+1, value=h)

yibu = '一支部'   # 一支部
erbu = '二支部'   # 二支部
biye = '已毕业'   # 已毕业

for i, name in enumerate(names_data):
    r = i + 2
    ws.cell(row=r, column=1, value=i+1)
    ws.cell(row=r, column=2, value=name)
    ws.cell(row=r, column=3, value=yibu if i < 5 else erbu)
    ws.cell(row=r, column=4, value=biye)

wb.save('test_data/test_party.xlsx')
print(f'Excel created: {len(names_data)} names')

# --- Verify ---
from openpyxl import load_workbook
wb2 = load_workbook('test_data/test_party.xlsx')
ws2 = wb2.active
cell_val = ws2.cell(row=2, column=2).value
header_val = ws2.cell(row=1, column=2).value
name_ok = cell_val == '张三'
print(f'Header B1: {repr(header_val)}')
print(f'First name: {repr(cell_val)}')
print(f'Verified: {"PASS" if name_ok else "FAIL"}')

# --- WeChat text ---
wechat_lines = [
    '群成员 (12)',           # 群成员 (12)
    '群主: 刘磊',       # 群主: 刘磊
    '张三-XX街道党建办',  # 张三-XX街道党建办
    '王小明(20级硕士)',      # 王小明(20级硕士)
    '欧阳锋',                # 欧阳锋
    '赵六六[一支部]',        # 赵六六[一支部]
    '周八—已改备注',    # 周八—已改备注
    '钱九',                      # 钱九
    '刘磊',                      # 刘磊
    '陈静~★',               # 陈静~★
    '测试人员A',         # 测试人员A
    '外来人员',          # 外来人员
    '孙七',                      # 孙七
]
with codecs.open('test_data/test_wechat_members.txt', 'w', 'utf-8') as f:
    f.write('\n'.join(wechat_lines))
print(f'WeChat text created: {len(wechat_lines)} lines')

# Verify text file
with codecs.open('test_data/test_wechat_members.txt', 'r', 'utf-8') as f:
    txt = f.read()
lines = txt.strip().split('\n')
print(f'TXT line 3: {repr(lines[2])}')
txt_ok = '张三' in txt  # 张三
print(f'Text verified: {"PASS" if txt_ok else "FAIL"}')

print('\nALL DONE - Ready for party_wechat_check test!')
