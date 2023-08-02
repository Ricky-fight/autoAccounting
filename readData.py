import xlwings as xl
import os
import re
path = r"."
os.chdir(path)

workbook = xl.Book('2023租金07月.xlsx')
sheet = workbook.sheets['合同']

rows = sheet.used_range.rows.count
cols = sheet.used_range.columns.count

dayOneCell = sheet[0, cols-31]

RENT = 0
DEPOSIT = 1
while True:
    prompt = input()
    if prompt == '#':
        break
    if '【押】' in prompt:
        mode = DEPOSIT
    else:
        mode = RENT
    prompt = re.sub(r'【.*】', '', prompt)
    prompt = re.sub(r'更正：', '', prompt)
    temp = prompt.split(' ')
    if len(temp) < 3:
        print('此命令不符合以下格式：“姓名 车牌号 微信/支付宝xxxx”，请重新输入')
        continue
    if re.match(r'沪?[A-Za-z0-9]{6,7}',temp[0]):
        plate = temp[0]
        name = temp[1]
    elif re.match(r'沪?[A-Za-z0-9]{6,7}',temp[1]):
        plate = temp[1]
        name = temp[0]
    print(name + ', ' + plate)
    for a in sheet['C2:C'+str(rows)]:
        if a.value == name:
            r = str(a.row)
            for b in sheet[f'A{r}:Z{r}']:
                tempPlate = re.sub(r'沪', plate)
                if b.value == tempPlate:
                    isPlateMatched = True
                if b.value == '生效中':
                    isContrastValid = True
            if not (isContrastValid and isPlateMatched):
                print('车牌号不匹配或合同不为生效中')
                continue
            # 匹配成功，开始操作记账
            # needPayment = sheet[f'U{a.row}']
            balance = sheet[f'W{a.row}']


