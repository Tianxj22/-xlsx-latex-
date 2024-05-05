#——————————将xls表格转化成latex的格式——————————#

#——————————————————导入文件——————————————————#
import openpyxl
import os
from file_and_path import *
from config_dealer import *
from xls_reader import *

config = get_config()
full_line = config["full_line"]

#———————————————————函数————————————————————#

def raw_value(cell, accuracy) -> str:
    '''获取单元格的原始数据并返回'''
    if type(cell.val) == float:
        return f"{round(cell.val, accuracy)}"
    elif cell.val == None:
        return config['empty_filler']
    else:
        return f"{cell.val}"

def to_fomular(origin: str) -> str:
    '''将单元格的值转化为LaTex的数学公式'''
    return f"${origin}$"

def to_multi(origin: str, cell: MergeCell) -> str:
    '''将单元格的值转化成对应的合并单元格的值'''
    if cell.pos[1] == cell.lt_pos[1]:
        return f"\\multicolumn{{{cell.shape[1]}}}{{{series_c(1, full_line)}}}{{\\multirow{{{cell.shape[0]}}}*{{{origin}}}}}"
    else:
        return ""

def add_tail(origin: str, cell, end=False):
    '''为单元格添加结尾的后缀'''
    if end:
        return f"{origin}\\\\\n"
    elif type(cell) == BaseCell or (type(cell) == MergeCell and cell.lt_pos[1] + cell.shape[1] - 1 == cell.pos[1]):
        return f"{origin} & "
    return origin

def series_c(num: int, full_line=False):
    '''返回好几个c, 如果要画满线就在中间穿插竖线'''
    rt = ""
    if full_line:
        rt += "|"
    return rt + ('c' + rt) * num

def single_value(cell, accuracy: 2, end=False)->str:
    '''将单个单元格的数据转化成字符格式
        cell: 单元格
        accuracy: 精确到小数点后几位
        end: 该单元格是否为末尾的单元格'''
    # 这里负责处理小数
    rt = raw_value(cell=cell, accuracy=accuracy)
    # 处理公式
    if config['formular']:
        rt = to_fomular(rt)
    # 处理合并的单元格
    if type(cell) == MergeCell:
        rt = to_multi(rt, cell)
    # 处理结尾
    rt = add_tail(rt, cell, end=end)
    return rt

def add_line(row: list, full_line = False) -> str:
    '''根据每一行的数据绘制出相应的线
        row: 一行的单元格数据
        full_line: 是否要划线'''
    if not full_line:
        return ""
    # 查看每一个单元格 如果是合并的单元格，且当前不是对应的最后一列，就绘制一条线
    rt = ""
    for idx in range(len(row)):
        if not (type(row[idx]) == MergeCell and row[idx].lt_pos[0] + row[idx].shape[0] - 1 != row[idx].pos[0]):
            rt += f" \\cline{{{idx + 1}-{idx + 1}}}"
    return rt

def xlsx_to_latex(filename: str, xlsx_dir: str,
    output_dir: str, accuracy = 2):
    '''将指定的某个名称的表格转化成latex代码,并输出到指定目录下,存储在同名的txt文件下'''
    reader = XlsReader(file_name=os.path.join(xlsx_dir, filename), data_only=True)
    f = open(os.path.join(output_dir, f"{filename.split('.')[0]}.txt"), 'w')
    for sheetname, data in reader.data.items():
        print(f"\t现在处理工作簿: {sheetname}")
        data_string = "\t\t\hline"
        n = max([len(row) for row in data])
        for row in data:
            data_string += "\t\t"
            for cell in row[:-1]:
                # 如果含有公式的话，开头结尾要加上$$
                data_string += f"{single_value(cell, accuracy, end=False)}"
            data_string += f"{single_value(row[-1], accuracy, end=True)} {add_line(row=row, full_line=full_line)}"
        f.write("\\begin{table}\n")
        f.write("\t\\centering\n")
        f.write("\t\\begin{tabular}{")
        f.write(f"{series_c(n, full_line)}")
        f.write("}\n")
        # 这里是文档的内容
        f.write(data_string)
        f.write("\t\\end{tabular}\n")
        f.write("\t\\caption{" + sheetname + "}\n")
        # f.write("\t\\lable{tab:" + sheetname + "}\n")
        f.write("\\end{table}\n\n\n")


#——————————————————主函数———————————————————#