import pandas as pd


def read_and_print_excel(file_name):
    # 尝试打开Excel文件并读取所有工作表的数据
    try:
        # 使用pandas的read_excel函数，将sheet_name设置为None以加载所有sheets
        xls = pd.read_excel(file_name, engine='openpyxl', sheet_name=None)
        print(f"Data from {file_name}:")

        # xls现在是一个字典，键是工作表名，值是对应的DataFrame
        for sheet_name, df in xls.items():
            print(f"Sheet name: {sheet_name}")
            print(df)  # 打印每个工作表DataFrame的前几行数据
            print("\n")  # 添加换行以便于区分不同工作表的输出

    except Exception as e:
        print(f"Failed to read {file_name}: {str(e)}")


# 文件路径
dir1 = '/Users/mac/PycharmProjects/train/汽车-比亚迪历史数据.xlsx'
dir2 = '/Users/mac/PycharmProjects/train/汽车-东风汽车历史数据.xlsx'

# 读取并打印每个文件的所有工作表内容
read_and_print_excel(dir1)
read_and_print_excel(dir2)
