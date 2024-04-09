#————————关于文件路径获取的代码——————————#

import os

def get_root()->str:
    '''获取代码所在的目录'''
    return os.path.dirname(os.path.abspath(__file__))

def get_file_list(dir_name: str):
    '''返还所给文件目录下的所有文件路径的列表和文件名的列表'''
    return [os.path.join(dir_name, file) for file in os.listdir(dir_name)], os.listdir(dir_name)
