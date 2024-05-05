#————————————————————————————————————————单个单元格的类————————————————————————————————————————#
#————————————————————————————————————————————常量————————————————————————————————————————————#



#———————————————————————————————————————————头文件————————————————————————————————————————————#

import os
import openpyxl

#—————————————————————————————————————————————类—————————————————————————————————————————————#

class BaseCell():
    '''最普通的单元格'''

    def __init__(self, val = None, pos = [0, 0]) -> None:
        '''初始化单元格
            val: 单元格的值
        '''
        self.val = val
        self.pos = pos

    def __str__(self) -> str:
        return f"val: {self.val}, pos: {self.pos}"

class MergeCell(BaseCell):
    '''合并的单元格'''

    def __init__(self, val=None, pos=[0, 0], lt_pos=[0, 0], shape=[1, 1]) -> None:
        '''初始化单元格
            val: 单元格的值
            lt_pos: 左上角的表格坐标
            shape: 合并的单元格的尺寸'''
        super().__init__(val, pos)
        self.lt_pos = lt_pos
        self.is_lt = (lt_pos == pos)
        self.shape = shape
    
    def __str__(self) -> str:
        return f"{super().__str__()}, lt_pos: {self.lt_pos}, is_lt: {self.is_lt}, shape: {self.shape}"

#————————————————————————————————————————————程序————————————————————————————————————————————#