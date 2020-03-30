# -*- coding: utf-8 -*-
"""
Created on Fri Dec 13 15:19:35 2019

@author: csr

openpyxl中 bounds元组定义的顺序是：0-3 min_col, min_row, max_col, max_row

表的配置数据的cbounds
根据第一列的单元格合并信息来确定单个检查区域
检查区域的内容包括 第一列单元格的Id 作为索引，检查区域的四个顶点
"""

from scr.Helper.CompareBounds import CompareBounds
from scr.Helper.SheetCopy import CopyCell
import scr.Helper.DebugHelper as DebugHelper


class ReadOnlySheetReader:
        
    def __init__(self,sheet,excel_title):
        self.excel_title = excel_title
        if sheet is None:
             raise Exception("【读取异常】Sheet 对象为None")
        self.sheet = sheet
        self._initBoundsInfo()
        
    #获取Title作用行
    def _getMaxTitleRow(self):
        #第一列从0向下遍历一直到第一个数字
        sheet = self.sheet
        title_row = 0
        for col_items in sheet.iter_cols(max_col = 1):
            for cell in col_items:
                if cell.value != None:
                    if  isinstance(cell.value,int):
                        break
                title_row += 1
        return title_row
                

    #获取当前表的有效宽度
    def _getConfigWidth(self):
        sheet = self.sheet
        max_value_column = 1
        #取第一行最后一个有有效值
        for row_items in sheet.iter_cols(max_row = 1):
            for cell in row_items:
                if cell.value != None:
                    if max_value_column < cell.column:
                        max_value_column = cell.column        
        #匹配最后一个有效值的合并单元格 并获取单元格的最大列
        for merge_cell in sheet.merged_cells.ranges:             
            if merge_cell.bounds[0] == max_value_column and merge_cell.bounds[1] == 1:
                if max_value_column < merge_cell.bounds[2]:
                    max_value_column = merge_cell.bounds[2]                    
        return max_value_column

    #初始化 每个sheet的检查区域Bounds信息
    def _initBoundsInfo(self): 
        self.max_title_row = self._getMaxTitleRow()
        self.compare_width = self._getConfigWidth()
        self.title_bounds = CompareBounds(1,1,self.compare_width,self.max_title_row,"Title")
        #表的配置数据的cbounds（初始化缓存对象）
        sheet = self.sheet
        body_bounds_dic =  dict()
        body_value_list = list()  
        for merge_cell in sheet.merged_cells.ranges: 
            #先第一列计算有合并单元格的的内容
            if self._isKeyMergeCell(merge_cell):
                key_cell = sheet.cell(merge_cell.bounds[1],merge_cell.bounds[0])
                cb =  CompareBounds(key_cell.column,key_cell.row,self.compare_width,merge_cell.bounds[3],key_cell.value)
                cb.SetMergeCellState(merge_cell)
                body_bounds_dic[key_cell.value] = cb
        for col_items in sheet.iter_cols(max_col = 1):
            #在计算计算第一列没有有合并单元格的单位
            for cell in col_items:
                if cell.row <= self.max_title_row:
                    continue
                if cell.value != None:
                    body_value_list.append(cell.value)
                    if cell.value  not in body_bounds_dic.keys():
                        body_bounds_dic[cell.value] = CompareBounds(cell.column,cell.row,self.compare_width,cell.row,cell.value)                                        
        
        self.body_value_list =  body_value_list
        self.body_bounds_dic = body_bounds_dic                                    

   #是否是cbounds的索引key的合并单元格
    def _isKeyMergeCell(self,merge_cell):                     
         result = False
         if merge_cell.bounds[0] == 1 and merge_cell.bounds[0] == merge_cell.bounds[2]:
             if merge_cell.bounds[1] > self.max_title_row and merge_cell.bounds[3] > self.max_title_row:
                 result = True
         return result
    
    #是否是Cbounds内的合并单元格
    def _isInCboundsMergeCell(self,merge_cell,cbounds,is_just_row = False):
        #0-3 min_col, min_row, max_col, max_row   
        min_col = merge_cell.bounds[0]        
        min_row = merge_cell.bounds[1]        
        max_col = merge_cell.bounds[2]        
        max_row = merge_cell.bounds[3]
        if  is_just_row:
            if  min_row >= cbounds.min_row  and max_row <= cbounds.max_row:
                return True
        else:
            if min_col >= cbounds.min_col and min_row >= cbounds.min_row and max_col <= cbounds.max_col and max_row <= cbounds.max_row:
                return True
        return False