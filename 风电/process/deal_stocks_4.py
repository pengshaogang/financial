import copy
import os
import openpyxl

from openpyxl.styles import Alignment
from openpyxl.styles import NamedStyle
from openpyxl.styles import PatternFill







file_names = [f for f in os.listdir('.') if f.startswith('风电')]


columns_to_delete = ['C', 'E', 'G', 'I', 'K']

for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name = [s for s in workbook.sheetnames if "资产负债表" in s]
    cal_sheet_name2 = [s for s in workbook.sheetnames if "利润表" in s]
    cal_sheet_name3 = [s for s in workbook.sheetnames if "现金流量表" in s]

    for sheet_name in cal_sheet_name:
        sheet_cal = workbook[sheet_name]
        first_row = list(sheet_cal[1])  # 行索引在 openpyxl 中是从1开始的，所以第二行是2
        # 遍历第二行的单元格
        for cell in first_row:
            if 'diff' in str(cell.value).lower() in str(cell.value).lower():  # 检查单元格中是否包含 'ave'
                # 删除包含 'ave' 的列
                sheet_cal.delete_cols(cell.column)  # cell.column 为该单元格所在列的列号

    for sheet_name2 in cal_sheet_name2:
        sheet_cal2 = workbook[sheet_name2]
        first_row2 = list(sheet_cal2[1])  # 行索引在 openpyxl 中是从1开始的，所以第二行是2
        # 遍历第二行的单元格
        for cell in first_row2:
            if 'diff' in str(cell.value).lower() in str(cell.value).lower():  # 检查单元格中是否包含 'ave'
                # 删除包含 'ave' 的列
                sheet_cal2.delete_cols(cell.column)  # cell.column 为该单元格所在列的列号

    for sheet_name3 in cal_sheet_name3:
        sheet_cal3 = workbook[sheet_name3]
        first_row3 = list(sheet_cal3[1])  # 行索引在 openpyxl 中是从1开始的，所以第二行是2
        # 遍历第二行的单元格
        for cell in first_row3:
            if 'diff' in str(cell.value).lower() in str(cell.value).lower():  # 检查单元格中是否包含 'ave'
                # 删除包含 'ave' 的列
                sheet_cal3.delete_cols(cell.column)  # cell.column 为该单元格所在列的列号

    # 保存修改
    workbook.save(file_name)

rows_to_process = []
rows_to_process.extend(range(2,100))

def subtract_lists(list1, list2):
    result = []
    for x, y in zip(list1, list2):
        try:
            # 首先检查x和y是否为None，如果为None，则直接跳过转换
            if x is None or y is None:
                result.append('--')
                continue  # 继续处理下一对元素
            # Check if both elements are numbers before subtracting
            if isinstance(float(x), (int, float)) and isinstance(float(y), (int, float)) and float(y) != 0:
                result.append((float(x) - float(y)) / abs(float(y)))
            else:
                result.append('--')  # Append '-' when either element is not numeric
        except ValueError:
            # 如果x或y不能被转换成浮点数，则进入这个分支
            result.append('--')  # 附加'-'表示无法进行数值运算
    return result


percent_style_name = 'percent_style'

# 检查样式是否已存在
existing_styles = [style for style in workbook.named_styles]
if percent_style_name not in existing_styles:
    percent_style = NamedStyle(name=percent_style_name, number_format='0.00%')
    workbook.add_named_style(percent_style)
else:
    percent_style = next(style for style in workbook.named_styles if style == percent_style_name)


right_alignment = Alignment(horizontal='right')

column_map = {
    'B': 'C',
    'C': 'E',
    'D': 'G',
    'E': 'I',
    'F': 'K',
}

columns_to_update = ['C', 'E', 'G', 'I', 'K']
columns_to_update_l = ['D', 'F', 'H', 'J', 'L']

red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')

for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name = [s for s in workbook.sheetnames if "资产负债表" in s]

    for sheet_name in cal_sheet_name:
        sheet_cal = workbook[sheet_name]

        for col in ['B', 'C', 'D', 'E', 'F', 'G']:
            current_values = []
            for row_num in rows_to_process:
                value = sheet_cal[f'{col}{row_num}'].value
                current_values.append(value)

            if col == 'B':
                values_B = current_values
            elif col == 'C':
                values_C = current_values
            elif col == 'D':
                values_D = current_values
            elif col == 'E':
                values_E = current_values
            elif col == 'F':
                values_F = current_values
            elif col == 'G':
                values_G = current_values

        B_C_diff = subtract_lists(values_B, values_C)
        C_D_diff = subtract_lists(values_C, values_D)
        D_E_diff = subtract_lists(values_D, values_E)
        E_F_diff = subtract_lists(values_E, values_F)
        F_G_diff = subtract_lists(values_F, values_G)
        # print(file_name)
        # print(values_B)
        # print(values_C)
        # print(D_E_diff)
        # print(E_F_diff)
        # print(F_G_diff)

        original_columns = ['F', 'E', 'D', 'C', 'B']
        for col in original_columns:
            col_index = openpyxl.utils.column_index_from_string(col) + 1
            sheet_cal.insert_cols(col_index)  # 插入新列

        for col1 in columns_to_update:
            new_col_index = openpyxl.utils.column_index_from_string(col1)
            # 设置第一行的对应列的值为 'diff'
            sheet_cal.cell(row=1, column=new_col_index).value = 'diff'

        for i, row_num1 in enumerate(rows_to_process):
            # 插入B_D_diff数据到D列
            # print(file_name)
            # print("valueB = ", B_D_diff[i])

            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('C'), value=B_C_diff[i])
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('E'), value=C_D_diff[i])

            # 插入F_H_diff数据到J列
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('G'), value=D_E_diff[i])

            # 插入H_J_diff数据到M列
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('I'), value=E_F_diff[i])
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('K'), value=F_G_diff[i])


        for row2 in rows_to_process:
            for col in ['C', 'E', 'G', 'I', 'K']:
                cell1 = sheet_cal[f'{col}{row2}']
                if cell1.value != '--':
                    cell1.style = percent_style

    workbook.save(file_name)  # 保存修改后的工作簿

for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name = [s for s in workbook.sheetnames if "资产负债表" in s]

    for sheet_name in cal_sheet_name:
        cal_sheet = workbook[sheet_name]

        for col in columns_to_update:
            col_index = openpyxl.utils.column_index_from_string(col)
            for row in rows_to_process:
                cell = cal_sheet.cell(row=row, column=col_index)
                if isinstance(cell.value, int) or isinstance(cell.value, float):
                    if abs(cell.value) >= 0.5:
                        cell.fill = red_fill
                    elif abs(cell.value) >= 0.25:
                        cell.fill = yellow_fill

    workbook.save(file_name)



for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_namel = [s for s in workbook.sheetnames if "利润表" in s]

    for sheet_namel in cal_sheet_namel:
        sheet_cal_l = workbook[sheet_namel]

        for col in ['C', 'D', 'E', 'F', 'G', 'H']:
            current_values_l = []
            for row_num in rows_to_process:
                value = sheet_cal_l[f'{col}{row_num}'].value
                current_values_l.append(value)

            if col == 'C':
                values_C_l = current_values_l
            elif col == 'D':
                values_D_l = current_values_l
            elif col == 'E':
                values_E_l = current_values_l
            elif col == 'F':
                values_F_l = current_values_l
            elif col == 'G':
                values_G_l = current_values_l
            elif col == 'H':
                values_H_l = current_values_l

        C_D_diff_l = subtract_lists(values_C_l, values_D_l)
        D_E_diff_l = subtract_lists(values_D_l, values_E_l)
        E_F_diff_l = subtract_lists(values_E_l, values_F_l)
        F_G_diff_l = subtract_lists(values_F_l, values_G_l)
        G_H_diff_l = subtract_lists(values_G_l, values_H_l)
        # print(file_name)
        # print(values_C_l)
        # print(values_D_l)
        # print(D_E_diff)
        # print(E_F_diff)
        # print(F_G_diff)

        original_columns = ['G', 'F', 'E', 'D', 'C']
        for col in original_columns:
            col_index = openpyxl.utils.column_index_from_string(col) + 1
            sheet_cal_l.insert_cols(col_index)  # 插入新列

        for col1 in columns_to_update_l:
            new_col_index_l = openpyxl.utils.column_index_from_string(col1)
            # 设置第一行的对应列的值为 'diff'
            sheet_cal_l.cell(row=1, column=new_col_index_l).value = 'diff'

        for i, row_num2 in enumerate(rows_to_process):
            # 插入B_D_diff数据到D列
            # print(file_name)
            # print("valueB = ", B_D_diff[i])

            sheet_cal_l.cell(row=row_num2, column=openpyxl.utils.column_index_from_string('D'), value=C_D_diff_l[i])
            sheet_cal_l.cell(row=row_num2, column=openpyxl.utils.column_index_from_string('F'), value=D_E_diff_l[i])

            # 插入F_H_diff数据到J列
            sheet_cal_l.cell(row=row_num2, column=openpyxl.utils.column_index_from_string('H'), value=E_F_diff_l[i])

            # 插入H_J_diff数据到M列
            sheet_cal_l.cell(row=row_num2, column=openpyxl.utils.column_index_from_string('J'), value=F_G_diff_l[i])
            sheet_cal_l.cell(row=row_num2, column=openpyxl.utils.column_index_from_string('L'), value=G_H_diff_l[i])


        for row3 in rows_to_process:
            for col in ['D', 'F', 'H', 'J', 'L']:
                cell2 = sheet_cal_l[f'{col}{row3}']
                if cell2.value != '--':
                    cell2.style = percent_style

    workbook.save(file_name)  # 保存修改后的工作簿

for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name_l = [s for s in workbook.sheetnames if "利润表" in s]

    for sheet_name in cal_sheet_name_l:
        cal_sheet = workbook[sheet_name]

        for col in columns_to_update_l:
            col_index = openpyxl.utils.column_index_from_string(col)
            for row in rows_to_process:
                cell = cal_sheet.cell(row=row, column=col_index)
                if isinstance(cell.value, int) or isinstance(cell.value, float):
                    if abs(cell.value) >= 0.5:
                        cell.fill = red_fill
                    elif abs(cell.value) >= 0.25:
                        cell.fill = yellow_fill

    workbook.save(file_name)


for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_namel_x = [s for s in workbook.sheetnames if "现金流量表" in s]

    for sheet_name_x in cal_sheet_namel_x:
        sheet_cal_x = workbook[sheet_name_x]

        for col in ['C', 'D', 'E', 'F', 'G', 'H']:
            current_values_x = []
            for row_num in rows_to_process:
                value = sheet_cal_x[f'{col}{row_num}'].value
                current_values_x.append(value)

            if col == 'C':
                values_C_x = current_values_x
            elif col == 'D':
                values_D_x = current_values_x
            elif col == 'E':
                values_E_x = current_values_x
            elif col == 'F':
                values_F_x = current_values_x
            elif col == 'G':
                values_G_x = current_values_x
            elif col == 'H':
                values_H_x = current_values_x

        C_D_diff_x = subtract_lists(values_C_x, values_D_x)
        D_E_diff_x = subtract_lists(values_D_x, values_E_x)
        E_F_diff_x = subtract_lists(values_E_x, values_F_x)
        F_G_diff_x = subtract_lists(values_F_x, values_G_x)
        G_H_diff_x = subtract_lists(values_G_x, values_H_x)
        # print(file_name)
        # print(values_C_x)
        # print(values_D_x)
        # print(D_E_diff)
        # print(E_F_diff)
        # print(F_G_diff)

        original_columns = ['G', 'F', 'E', 'D', 'C']
        for col in original_columns:
            col_index_x = openpyxl.utils.column_index_from_string(col) + 1
            sheet_cal_x.insert_cols(col_index_x)  # 插入新列

        for col2 in columns_to_update_l:
            new_col_index_x = openpyxl.utils.column_index_from_string(col2)
            # 设置第一行的对应列的值为 'diff'
            sheet_cal_x.cell(row=1, column=new_col_index_x).value = 'diff'

        for i, row_num3 in enumerate(rows_to_process):
            # 插入B_D_diff数据到D列
            # print(file_name)
            # print("valueB = ", B_D_diff[i])

            sheet_cal_x.cell(row=row_num3, column=openpyxl.utils.column_index_from_string('D'), value=C_D_diff_x[i])
            sheet_cal_x.cell(row=row_num3, column=openpyxl.utils.column_index_from_string('F'), value=D_E_diff_x[i])

            # 插入F_H_diff数据到J列
            sheet_cal_x.cell(row=row_num3, column=openpyxl.utils.column_index_from_string('H'), value=E_F_diff_x[i])

            # 插入H_J_diff数据到M列
            sheet_cal_x.cell(row=row_num3, column=openpyxl.utils.column_index_from_string('J'), value=F_G_diff_x[i])
            sheet_cal_x.cell(row=row_num3, column=openpyxl.utils.column_index_from_string('L'), value=G_H_diff_x[i])


        for row4 in rows_to_process:
            for col in ['D', 'F', 'H', 'J', 'L']:
                cell3 = sheet_cal_x[f'{col}{row4}']
                if cell3.value != '--':
                    cell3.style = percent_style

    workbook.save(file_name)  # 保存修改后的工作簿

for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name_x = [s for s in workbook.sheetnames if "现金流量表" in s]

    for sheet_name in cal_sheet_name_x:
        cal_sheet = workbook[sheet_name]

        for col in columns_to_update_l:
            col_index = openpyxl.utils.column_index_from_string(col)
            for row in rows_to_process:
                cell = cal_sheet.cell(row=row, column=col_index)
                if isinstance(cell.value, int) or isinstance(cell.value, float):
                    if abs(cell.value) >= 0.5:
                        cell.fill = red_fill
                    elif abs(cell.value) >= 0.25:
                        cell.fill = yellow_fill

    workbook.save(file_name)

