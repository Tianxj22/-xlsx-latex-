#——————————将xls表格转化成latex的格式——————————#

#——————————————————导入文件——————————————————#
import openpyxl
import os
from file_and_path import *
from config_dealer import *
from xls_reader import *

config = get_config()

#———————————————————函数————————————————————#

def is_fomular(origin: str) -> str:
    return f"${origin}$"



def single_value(cell, accuracy: 2)->str:
    '''将单个单元格的数据转化成字符格式'''
    rt = ''
    # 首先查看是否要写公式
    if config['formular']:
        rt += '$'

    # 这里负责处理合并单元格
    # if type(cell) == MergeCell:
    #     if cell.is_lt:


    # 这里负责处理小数
    if type(cell.val) == float:
        rt += f"{round(cell.val, accuracy)}"
    elif cell.val == None:
        rt += config['empty_filler']
    else:
        rt += f"{cell.val}"
    
    if config['formular']:
        rt += '$'
    return rt

def xlsx_to_latex(filename: str, xlsx_dir: str,
    output_dir: str, accuracy = 2):
    '''将指定的某个名称的表格转化成latex代码,并输出到指定目录下,存储在同名的txt文件下'''
    reader = XlsReader(file_name=os.path.join(xlsx_dir, filename), data_only=True)
    f = open(os.path.join(output_dir, f"{filename.split('.')[0]}.txt"), 'w')
    for sheetname, data in reader.data.items():
        print(f"\t现在处理工作簿: {sheetname}")
        data_string = ""
        n = 0
        for row in data:
            n = max(len(row), n)
            data_string += "\t\t"
            for cell in row[:-1]:
                # 如果含有公式的话，开头结尾要加上$$
                data_string += f"{single_value(cell, accuracy)} & "
            data_string += f"{single_value(cell, accuracy)}\\\\\n"

        # 下面是负责写文件的部分
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


#——————————————————主函数———————————————————#

cur_dir = get_root()
xlsx_dir = os.path.join(cur_dir, 'xlsx_file')
output_dir = os.path.join(cur_dir, 'output_file')
filenames = os.listdir(xlsx_dir)

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
