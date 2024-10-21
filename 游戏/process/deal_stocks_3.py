import copy
import os
import openpyxl

from openpyxl.styles import Alignment
from openpyxl.styles import NamedStyle
from openpyxl.styles import PatternFill







file_names = [f for f in os.listdir('.') if f.startswith('游戏')]


columns_to_delete = ['D', 'G', 'J', 'M', 'P']

for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name = [s for s in workbook.sheetnames if "计算公式" in s]

    for sheet_name in cal_sheet_name:
        sheet_cal = workbook[sheet_name]
        first_row = list(sheet_cal[1])  # 行索引在 openpyxl 中是从1开始的，所以第二行是2
        # 遍历第二行的单元格
        for cell in first_row:
            if 'diff' in str(cell.value).lower() in str(cell.value).lower():  # 检查单元格中是否包含 'ave'
                # 删除包含 'ave' 的列
                sheet_cal.delete_cols(cell.column)  # cell.column 为该单元格所在列的列号

    # 保存修改
    workbook.save(file_name)




values_b_1, values_c_1, values_d_1, values_e_1, values_f_1 = [], [], [], [], []#1代表第一个指标，b,c,d,e,f代表原先的单元格

rows_to_process = [2,3,4,5, 7,8]

rows_to_process.extend(range(10, 27))

rows_to_process.extend(range(31,39))

rows_to_process.extend(range(44,48))

rows_to_process.extend(range(63,65))
rows_to_process.extend(range(67,68))
rows_to_process.extend(range(69,70))
rows_to_process.extend(range(76,78))

columns = ['B', 'C', 'D', 'E', 'F']
keys = ['val1', 'val2', 'val3','val4', 'val6','val7']

keys.extend([f'val{i}' for i in range(9, 26)])  # 使用列表推导来生成和添加键
keys.extend([f'val{i}' for i in range(30,38)])
keys.extend([f'val{i}' for i in range(43,47)])
keys.extend([f'val{i}' for i in range(62,64)])
keys.extend([f'val{i}' for i in range(66,67)])
keys.extend([f'val{i}' for i in range(68,69)])
keys.extend([f'val{i}' for i in range(75,77)])


# 使用字典推导和循环初始化values字典
values = {col: {key: [] for key in keys} for col in columns}

def subtract_lists(list1, list2):
    result = []
    for x, y in zip(list1, list2):
        # Check if both elements are numbers before subtracting
        if isinstance(x, (int, float)) and isinstance(y, (int, float)):
            result.append((x - y) / abs(y))
        else:
            result.append('-')  # Append '-' when either element is not numeric
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

#保存到.xlsx

column_map = {
    'C': 'D',
    'E': 'G',
    'G': 'J',
    'I': 'M',
    'K': 'P'
}

columns_to_update = ['D', 'G', 'J', 'M', 'P']

red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')

for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name = [s for s in workbook.sheetnames if "计算公式" in s]

    for sheet_name in cal_sheet_name:
        sheet_cal = workbook[sheet_name]

        # 读取B2单元格的值并添加到列表中
        # 读取B2到F2单元格的值并添加到相应的列表中

        for col in ['B', 'D', 'F', 'H', 'J']:
            current_values = []
            for row_num in rows_to_process:
                value = sheet_cal[f'{col}{row_num}'].value
                current_values.append(value)

            if col == 'B':
                values_B = current_values
            elif col == 'D':
                values_D = current_values
            elif col == 'F':
                values_F = current_values
            elif col == 'H':
                values_H = current_values
            elif col == 'J':
                values_J = current_values

        # print(file_name)
        B_D_diff = subtract_lists(values_B, values_D)
        D_F_diff = subtract_lists(values_D, values_F)
        F_H_diff = subtract_lists(values_F, values_H)
        H_J_diff = subtract_lists(values_H, values_J)
        J_L_diff = 41 * ['-']
        # print(H_J_diff)

        original_columns = ['K', 'I', 'G', 'E', 'C']
        for col in original_columns:
            col_index = openpyxl.utils.column_index_from_string(col) + 1
            sheet_cal.insert_cols(col_index)  # 插入新列

        for col1 in columns_to_update:
            new_col_index = openpyxl.utils.column_index_from_string(col1)
            # 设置第一行的对应列的值为 'diff'
            sheet_cal.cell(row=1, column=new_col_index).value = 'diff'
            sheet_cal.cell(row=30, column=new_col_index).value = 'diff'
            sheet_cal.cell(row=43, column=new_col_index).value = 'diff'
            sheet_cal.cell(row=62, column=new_col_index).value = 'diff'
            sheet_cal.cell(row=66, column=new_col_index).value = 'diff'
            sheet_cal.cell(row=75, column=new_col_index).value = 'diff'

        for i, row_num1 in enumerate(rows_to_process):
            # 插入B_D_diff数据到D列
            # print(file_name)
            # print("valueB = ", B_D_diff[i])

            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('D'), value=B_D_diff[i])
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('G'), value=D_F_diff[i])

            # 插入F_H_diff数据到J列
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('J'), value=F_H_diff[i])

            # 插入H_J_diff数据到M列
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('M'), value=H_J_diff[i])
            sheet_cal.cell(row=row_num1, column=openpyxl.utils.column_index_from_string('P'), value=J_L_diff[i])


        sheet_cal.column_dimensions['D'].width = 20
        sheet_cal.column_dimensions['G'].width = 20
        sheet_cal.column_dimensions['J'].width = 20
        sheet_cal.column_dimensions['M'].width = 20
        sheet_cal.column_dimensions['P'].width = 20

        sheet_cal.column_dimensions['L'].width = 20
        sheet_cal.column_dimensions['N'].width = 20
        sheet_cal.column_dimensions['O'].width = 20

        for row2 in rows_to_process:
            for col in ['D', 'G', 'J', 'M', 'P']:
                cell1 = sheet_cal[f'{col}{row2}']
                if cell1.value != '-':
                    cell1.style = percent_style

        cells_to_align_right = ['D1', 'G1', 'J1', 'M1', 'P1',
                                'D2', 'G2', 'J2', 'M2', 'P2',
                                'D3', 'G3', 'J3', 'M3', 'P3',
                                'D4', 'G4', 'J4', 'M4', 'P4',
                                'D5', 'G5', 'J5', 'M5', 'P5',
                                'D7', 'G7', 'J7', 'M7', 'P7',
                                'D8', 'G8', 'J8', 'M8', 'P8',
                                'D9', 'G9', 'J9', 'M9', 'P9',
                                'D10', 'G10', 'J10', 'M10', 'P10',
                                'D11', 'G11', 'J11', 'M11', 'P11',
                                'D12', 'G12', 'J12', 'M12', 'P12',
                                'D13', 'G13', 'J13', 'M13', 'P13',
                                'D14', 'G14', 'J14', 'M14', 'P14',
                                'D15', 'G15', 'J15', 'M15', 'P15',
                                'D16', 'G16', 'J16', 'M16', 'P16',
                                'D17', 'G17', 'J17', 'M17', 'P17',
                                'D18', 'G18', 'J18', 'M18', 'P18',
                                'D19', 'G19', 'J19', 'M19', 'P19',
                                'D20', 'G20', 'J20', 'M20', 'P20',
                                'D21', 'G21', 'J21', 'M21', 'P21',
                                'D22', 'G22', 'J22', 'M22', 'P22',
                                'D23', 'G23', 'J23', 'M23', 'P23',
                                'D24', 'G24', 'J24', 'M24', 'P24',
                                'D25', 'G25', 'J25', 'M25', 'P25',
                                'D26', 'G26', 'J26', 'M26', 'P26',
                                'D30', 'G30', 'J30', 'M30', 'P30',
                                'D31', 'G31', 'J31', 'M31', 'P31',
                                'D32', 'G32', 'J32', 'M32', 'P32',
                                'D33', 'G33', 'J33', 'M33', 'P33',
                                'D34', 'G34', 'J34', 'M34', 'P34',
                                'D35', 'G35', 'J35', 'M35', 'P35',
                                'D36', 'G36', 'J36', 'M36', 'P36',
                                'D37', 'G37', 'J37', 'M37', 'P37',
                                'D38', 'G38', 'J38', 'M38', 'P38',
                                'D43', 'G43', 'J43', 'M43', 'P43',
                                'D44', 'G44', 'J44', 'M44', 'P44',
                                'D45', 'G45', 'J45', 'M45', 'P45',
                                'D46', 'G46', 'J46', 'M46', 'P46',
                                'D47', 'G47', 'J47', 'M47', 'P47',
                                'D62', 'G62', 'J62', 'M62', 'P62',
                                'D63', 'G63', 'J63', 'M63', 'P63',
                                'D64', 'G64', 'J64', 'M64', 'P64',
                                'B66', 'D66', 'F66', 'H66', 'J66',
                                'D66', 'G66', 'J66', 'M66', 'P66',
                                'D67', 'G67', 'J67', 'M67', 'P67',
                                'D69', 'G69', 'J69', 'M69', 'P69',
                                'D75', 'G75', 'J75', 'M75', 'P75',
                                'D76', 'G76', 'J76', 'M76', 'P76',
                                'D77', 'G77', 'J77', 'M77', 'P77'
                                ]

        for cell in cells_to_align_right:
                sheet_cal[cell].alignment = right_alignment



    workbook.save(file_name)  # 保存修改后的工作簿




for file_name in file_names:
    workbook = openpyxl.load_workbook(file_name)
    cal_sheet_name = [s for s in workbook.sheetnames if "计算公式" in s]

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




