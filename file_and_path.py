#————————关于文件路径获取的代码——————————#

import os

def get_root()->str:
    '''获取代码所在的目录'''
    return os.path.dirname(os.path.abspath(__file__))

def get_file_list(dir_name: str):
    '''返还所给文件目录下的所有文件路径的列表和文件名的列表'''
    return [os.path.join(dir_name, file) for file in os.listdir(dir_name)], os.listdir(dir_name)

# 获取脚本所在的路径
# filepath = os.path.abspath(__file__)
# print(filepath)
# filedir = os.path.dirname(filepath)
# print(filedir)
# # 获取表格文件所在的目录
# xlsx_dir = os.path.join(filedir, "xlsx_file")
# print(xlsx_dir)
# # 获取要输出的txt文件所在的目录
# output_dir = os.path.join(filedir, "output_file")
# print(output_dir)
# # 获取表格文件夹下所有的文件
# xlsx_files = [os.path.join(xlsx_dir, file) for file in os.listdir(xlsx_dir)]
# print(xlsx_files)

# print(get_file_list(os.path.join(get_root(), 'xlsx_file')))