#——————将xls表格转化成latex的格式—————#

import openpyxl
import os
from file_and_path import *

def xlsx_to_latex(filename: str, xlsx_dir: str, output_dir: str):
    '''将指定的某个名称的表格转化成latex代码,并输出到指定目录下,存储在同名的txt文件下'''
    workbook = openpyxl.load_workbook(filename=os.path.join(xlsx_dir, filename), data_only=True)
    f = open(os.path.join(output_dir, f"{filename.split('.')[0]}.txt"), 'w')
    for sheetname in workbook.sheetnames:
        worksheet = workbook[sheetname]
        print(f"\t现在处理工作簿: {sheetname}")
        data_string = ""
        n = 0
        for row in worksheet.rows:
            n = max(len(row), n)
            data_string += "\t\t"
            for cell in row[:-1]:
                data_string += f"{cell.value} & "
            data_string += f"{row[-1].value}\\\\\n"
        f.write("\\begin{table}\n")
        f.write("\t\\centering\n")
        f.write("\t\\begin{tabular}{")
        f.write("c" * n)
        f.write("}\n")
        # 这里是文档的内容
        f.write(data_string)
        f.write("\t\\end{tabular}\n")
        f.write("\t\\caption{" + sheetname + "}\n")
        # f.write("\t\\lable{tab:" + sheetname + "}\n")
        f.write("\\end{table}\n\n\n")

# workbook = openpyxl.load_workbook(filename=filename, data_only=True)
# print(workbook.worksheets)
# xlsx_to_latex(filename, xlsx_dir, output_dir)

cur_dir = get_root()
xlsx_dir = os.path.join(cur_dir, 'xlsx_file')
output_dir = os.path.join(cur_dir, 'output_file')
# print(xlsx_dir, output_dir)
filenames = os.listdir(xlsx_dir)
# print(filenames)

for i in range(len(filenames)):
    print(f"{i}: {filenames[i]}")

idx_list_str = input("请输入要处理的表格前面的标号(用空格隔开)，若是-1则全部进行处理: ")

# 下面要处理的文件名列表
operate_filenames = []
if idx_list_str == '-1':
    operate_filenames = filenames.copy()
else:
    for idx_str in idx_list_str.split(" "):
        try:
            operate_filenames.append(filenames[int(idx_str)])
        except:
            print(f"下标{idx_str}越界！")

print(operate_filenames)
for operate_filename in operate_filenames:
    print(f"现在转换表格: {operate_filename.split('.')[0]}")
    xlsx_to_latex(operate_filename, xlsx_dir, output_dir)