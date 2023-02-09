import openpyxl
import csv
import os

list = os.listdir(os.getcwd())
for csvfile in list:
    if '.csv' in csvfile:
        break
filenpathame_csv = csvfile
filenpathame_xlsx = 'metersphere导入文件.xlsx'
sheetname = 'Sheet'
column_list = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
               "V", "W", "X", "Y", "Z"]
case_status = "Completed"


def csv2xlsx():  # csv转为xlsx excel2007以上版本
    with open(filenpathame_csv, 'r', encoding='utf-8') as f:
        read = csv.reader(f)
        wb = openpyxl.Workbook()
        ws = wb.active
        for line in read:
            ws.append(line)
        wb.save(filenpathame_xlsx)  # 保存Excel


def fill_elsx():
    wb = openpyxl.load_workbook(filenpathame_xlsx)
    ws = wb[sheetname]
    rows = ws.max_row
    columns = ws.max_column
    ws["A1"] = "所属模块"
    ws["B1"] = "废弃2"
    ws["C1"] = "用例名称"
    ws["D1"] = "废弃"
    ws["E1"] = "废弃"
    ws["F1"] = "用例等级"
    ws["G1"] = "前置条件"
    ws["H1"] = "废弃"
    ws["I1"] = "废弃"
    ws["J1"] = "步骤描述"
    ws["K1"] = "预期结果"
    ws["L1"] = "废弃"
    ws["M1"] = "用例状态"
    ws["N1"] = "废弃14"
    ws["O1"] = "废弃15"
    ws["P1"] = "用例状态"

    ws.delete_cols(2)
    ws.delete_cols(3, 2)
    ws.delete_cols(5, 2)
    ws.delete_cols(7, 4)
    ws.delete_cols(8)
    for i in range(2, rows + 1):
        ws.cell(row=i, column=7, value=case_status)
        ws.cell(row=i, column=5, value=format_numbers_with_brackets(remove_new_lines(ws.cell(row=i, column=5).value)))
        ws.cell(row=i, column=6, value=format_numbers_with_brackets(remove_new_lines(ws.cell(row=i, column=6).value)))
    print("转换完成")
    wb.save(filenpathame_xlsx)


def format_numbers_with_brackets(string):
    # 用于更改步骤和预期结果的格式
    result = ""
    for i in range(len(string)):
        char = string[i]
        if char.isdigit():
            if i > 0 and string[i - 1] == "[" and i < len(string) - 1 and string[i + 1] == "]":
                if i - 1 != 0:
                    result += f"\n"
                else:
                    result += ""
                i += 1
            else:
                result += char
        elif char == "[" and string[i + 1].isdigit() or char == "]" and string[i - 1].isdigit():
            continue
        else:
            result += char
    return result


def remove_new_lines(string):
    return string.replace("\n", "")


if __name__ == '__main__':
    csv2xlsx()
    fill_elsx()

# Press the green button in the gutter to run the script.
