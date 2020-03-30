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


class SheetReader:
        
    def __init__(self,sheet,dataonly_sheet,excel_title):
        self.excel_title = excel_title
        if sheet is None:
             raise Exception("【读取异常】Sheet 对象为None")
        self.sheet = sheet
        self.dataonly_sheet = dataonly_sheet
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

    #更新表的数据cbound范围（不修改缓存对象）
    def UpdateBodyBounds(self):
        sheet = self.sheet
        cache = []
        for merge_cell in sheet.merged_cells.ranges: 
            #先第一列计算有合并单元格的的内容
            if self._isKeyMergeCell(merge_cell):
                key_cell = sheet.cell(merge_cell.bounds[1],merge_cell.bounds[0])
                cbounds = self.body_bounds_dic[key_cell.value]
                if cbounds:
                    cache.append(key_cell.value)
                    cbounds.SetMergeCellState(merge_cell)
                    cbounds.UpdateBounds(key_cell.column,key_cell.row,self.compare_width,merge_cell.bounds[3])
        for col_items in sheet.iter_cols(max_col = 1):
            #在计算计算第一列没有有合并单元格的单位
            for cell in col_items:
                if cell.row <= self.max_title_row:
                    continue
                if cell.value != None:
                    if cell.value not in cache:
                        cbounds = self.body_bounds_dic[cell.value]
                        if cbounds:
                            cbounds.UpdateBounds(cell.column,cell.row,self.compare_width,cell.row)
        cache = None


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

    #取目标区域的单元格数据和合并单元格
    def ApplyDiff2BodyBounds(self,mine_key_diffs,their_key_diffs):
        #SVN的合并机制是将their与base的差异合入mine
        self._writeDiff2SheetByKeyDiffs(their_key_diffs)

    #根据key_diffs写入merge
    def _writeDiff2SheetByKeyDiffs(self,key_diffs):
        target_reader = key_diffs["target_sheet_reader"]
        add_key_set = key_diffs["add_key"]
        del_key_set = key_diffs["del_key"]
        mod_key_set = key_diffs["mod_key"]
        self._tryWriteAdd(add_key_set,target_reader)
        self._tryWriteDel(del_key_set,target_reader)
        self._tryWriteMod(mod_key_set,target_reader)
                
    #写入新增
    def _tryWriteAdd(self,add_key_set,target_reader):
        if len(add_key_set) > 0:
            for key in add_key_set:
                cbounds = target_reader.body_bounds_dic[key]
                DebugHelper.Log("【新增】",cbounds.tostring())
                seat_key = self._seachCommonKeyUp(key,target_reader)
                new_Cb = self._insertCBounds(seat_key,cbounds)
                bounds_value = self._getCellInfosByCbounds(target_reader.sheet,cbounds)
                self._drawMergeCellsByCbounds(bounds_value["merge_cell"],new_Cb,has_key_merge_cell =True)
                self._writeSheetBoundsValue(key,bounds_value) 
    #写入删除
    def _tryWriteDel(self,del_key_set,target_reader):
        if len(del_key_set) > 0:
            for key in del_key_set:
                #先拆分所有的单元格
                self_cbounds = self.body_bounds_dic[key]
                DebugHelper.Log("【删除】",key,self_cbounds.tostring())
                self._cleanMergeCellsByCbounds(self_cbounds,has_key_merge_cell =True)
                self.body_value_list.remove(key)
                bounds_value = {"rect":[0]} 
                self._adapterCBounds(self.sheet,key,bounds_value)

    #写入修改
    def _tryWriteMod(self,mod_key_set,target_reader):
        if len(mod_key_set) > 0:
            for key in mod_key_set:
                #TODO 当修改的单位检查取得行由多行变更为一行时，可能会有问题
                #从目标sheet取信息
                cbounds = target_reader.body_bounds_dic[key]
                bounds_value = self._getCellInfosByCbounds(target_reader.sheet,cbounds)

                #先拆分所有的单元格
                self_cbounds = self.body_bounds_dic[key]
                self._cleanMergeCellsByCbounds(self_cbounds)

                self._adapterCBounds(self.sheet,key,bounds_value)
                DebugHelper.Log("【修改】",self_cbounds.tostring())
                self._drawMergeCellsByCbounds(bounds_value["merge_cell"],self_cbounds)
                self._writeSheetBoundsValue(key,bounds_value)


    #插入CBounds内相应数量的行
    def _insertCBounds(self,seat_key,CBounds):
        sheet = self.sheet
        key = CBounds.key
        diff_row = CBounds.max_row - CBounds.min_row + 1
        self_bounds = self.body_bounds_dic[seat_key]
        min_col = self_bounds.min_col
        max_col = self_bounds.max_col
        min_row = self_bounds.max_row + 1
        max_row = self_bounds.max_row + diff_row
        new_Cb = CompareBounds(min_col,min_row,max_col,max_row,key)

        index = self.body_value_list.index(seat_key)
        index += 1        
        self.body_value_list.insert(index,key)
        self.body_bounds_dic[key] = new_Cb
        #新增相应数量的行
        self.sheet.insert_rows(self_bounds.max_row + 1,diff_row)

        for merge_cell in sheet.merged_cells.ranges:
            if merge_cell.min_row > self_bounds.max_row:
                    merge_cell.min_row += diff_row
                    merge_cell.max_row += diff_row
        self.UpdateBodyBounds()
        return new_Cb

    #计算增删的行数 来处理对即将写入区域的行适配
    def _adapterCBounds(self,sheet,key,bounds_value):
        self_bounds = self.body_bounds_dic[key]
        heigh = self_bounds.max_row - self_bounds.min_row + 1
        height_diff = heigh - bounds_value["rect"][0]
        if height_diff > 0:
            #大于0就删除行
            self._updateMergeCellRangeDown(sheet,self_bounds,height_diff)
            start_row = self_bounds.max_row
            start_index = start_row - height_diff
            start_index += 1
            sheet.delete_rows(start_index,height_diff)
            self.UpdateBodyBounds()
        elif height_diff < 0:
            start_row = self_bounds.max_row
            #小于0就新增行
            self._updateMergeCellRangeDown(sheet,self_bounds,height_diff)
            start_index = start_row
            height_diff = - height_diff
            sheet.insert_rows(start_index,height_diff)
            self.UpdateBodyBounds()


    #向下更新合并单元格的位置
    def _updateMergeCellRangeDown(self,sheet,self_bounds,height_diff):
        sheet = self.sheet
        self_max_row = self_bounds.max_row
        if height_diff > 0:
            #删除先处理当前 在处理合并单元格
            if self_bounds.merge_cell:
                merge_cell = self_bounds.merge_cell
                heigh = self_bounds.max_row - self_bounds.min_row + 1
                if heigh > height_diff:
                    merge_cell.max_row -= height_diff
            for merge_cell in sheet.merged_cells.ranges:
                if merge_cell.min_row > self_max_row:
                        merge_cell.min_row -= height_diff
                        merge_cell.max_row -= height_diff
        elif height_diff < 0:
            #新增先处理合并单元格 在处理当前
            for merge_cell in sheet.merged_cells.ranges:
                if merge_cell.min_row > self_max_row:
                        merge_cell.min_row -= height_diff
                        merge_cell.max_row -= height_diff
            if self_bounds.merge_cell:
                self_bounds.merge_cell.max_row -= height_diff

    #向上搜寻共同存在的key
    def _seachCommonKeyUp(self,key,sheet_reader):
        index = sheet_reader.body_value_list.index(key)
        vlu_list = sheet_reader.body_value_list[0:index + 1]
        vlu_list.reverse()
        for key in vlu_list:
            if key in self.body_value_list:
                return key
        #默认最后一个
        return self.body_value_list[-1]
    
    def _writeSheetBoundsValue(self,key,bounds_value):
        sheet = self.sheet
        self_bounds = self.body_bounds_dic[key]
        buffer_cells = bounds_value['cells'][:]
        #开始写入
        for col_items in sheet.iter_rows(self_bounds.min_row,self_bounds.max_row,self_bounds.min_col,self_bounds.max_col):
            buffer_col_items = buffer_cells.pop(0)
            for cell in col_items:
                source_cell = buffer_col_items.pop(0)
                #不处理只读的合并单元格
                CopyCell(cell,source_cell)                

        #获取cbounds内的写入需要的数据
    def _getCellInfosByCbounds(self,sheet,cbounds):
        #bounds_value的cell坐标都是相对坐标
        bounds_value = {}
        bounds_value['cells'] = []
        bounds_value['merge_cell'] = []
        bounds_value['rect'] = (cbounds.max_row - cbounds.min_row + 1,cbounds.max_col - cbounds.min_col + 1)
        for col_items in sheet.iter_rows(cbounds.min_row,cbounds.max_row,cbounds.min_col,cbounds.max_col):
            col_list = []
            for cell in col_items:
                col_list.append(cell)
            bounds_value['cells'].append(col_list)

        #范围内的合并单元格
        for merge_cell in sheet.merged_cells.ranges:
            if self._isInCboundsMergeCell(merge_cell,cbounds):
                #若该范围内有合并单元格，则定点位置设置成相对位置
                #0-3 min_col, min_row, max_col, max_row   
                min_row = merge_cell.bounds[1]        
                max_row = merge_cell.bounds[3]

                offset_min_col = merge_cell.min_col
                offset_min_row = min_row - cbounds.min_row
                offset_max_col = merge_cell.max_col
                offset_max_row = max_row - cbounds.min_row
                merge_cbounds = CompareBounds(offset_min_col,offset_min_row,offset_max_col,offset_max_row)
                bounds_value['merge_cell'].append(merge_cbounds)
        return bounds_value

    #清除Cbounds内的合并单元格
    def _cleanMergeCellsByCbounds(self,cbounds,has_key_merge_cell = False):
        sheet = self.sheet
        unmerge_cells = []
        for merge_cell in sheet.merged_cells.ranges:
            #非数据区域得合并单元格也需要处理
            if self._isInCboundsMergeCell(merge_cell,cbounds,is_just_row = True):
                if self._isKeyMergeCell(merge_cell):
                    if has_key_merge_cell:
                        unmerge_cells.append(merge_cell)
                else:
                    unmerge_cells.append(merge_cell)
        
        for merge_cell in unmerge_cells:
            sheet.unmerge_cells(start_row = merge_cell.min_row,start_column = merge_cell.min_col,end_row = merge_cell.max_row,end_column = merge_cell.max_col)

    def _drawMergeCellsByCbounds(self,merge_cell_cbounds,cbounds,has_key_merge_cell = False):
        sheet = self.sheet
        for cell_bounds in merge_cell_cbounds:
            need_merge = False            
            #是否是cbounds的索引key的合并单元格
            if cell_bounds.min_col == 1 and cell_bounds.max_col == 1:
                if has_key_merge_cell:
                    need_merge = True
            else:
                need_merge = True
            if need_merge:
                cell_bounds.min_row += cbounds.min_row
                cell_bounds.max_row += cbounds.min_row
                sheet.merge_cells(start_row = cell_bounds.min_row,start_column = cell_bounds.min_col,end_row = cell_bounds.max_row,end_column = cell_bounds.max_col)
        

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

