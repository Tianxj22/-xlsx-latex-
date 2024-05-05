#————————————————————————————程序运行的主程序——————————————————————————————————————#

#————————————————————————————————头文件——————————————————————————————————————————#

from xls_to_latex import *


#——————————————————————————————————程序——————————————————————————————————————————#



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

print(f"处理表格列表为: {operate_filenames}")
for operate_filename in operate_filenames:
    print(f"现在转换表格: {operate_filename.split('.')[0]}")
    xlsx_to_latex(operate_filename, xlsx_dir, output_dir)