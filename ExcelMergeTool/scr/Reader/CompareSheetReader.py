# -*- coding: utf-8 -*-
"""
Created on Sat Dec 14 17:47:11 2019

@author: admin

将每个cell的value的值拼接成一个字符串 并存储进行list比较
以收集的CompareBounds为单位进行新增 删除 修改

以下情况的合并待定（不处理或 不支持该内容合并）：
1.their和mine有一方修改了表头数据
2.their和mine都修改相同索引的bounds
3.新增的Sheet不处理（受制于openxl的复制sheet操作不能在两个excel之间操作）
4.公式不处理
"""

import scr.Helper.DebugHelper as DebugHelper

#判断是否有中文
def IsHasChinese(check_str):
    for ch in check_str:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False

class CompareSheetReader:
    
    def __init__(self):
        self.value_diff = dict()

    #表头是否有变更
    def IsHasTitleDiff(self,target_sheet_reader,base_sheet_reader):
        target_bounds = target_sheet_reader.title_bounds
        base_bounds = base_sheet_reader.title_bounds
        #作用域差异
        if target_bounds.bounds != base_bounds.bounds:
            return True
        #合并单元格变更
        vlu_diff_1 = self.HasMergeCellsDiff(None,target_sheet_reader,base_sheet_reader)
        vlu_diff_2 = self.HasMergeCellsDiff(None,target_sheet_reader,base_sheet_reader)
        if vlu_diff_1 or vlu_diff_2:
            return True
        #批注变更
        vlu_diff_1 = self.HasContentsDiff(None,target_sheet_reader,base_sheet_reader)
        vlu_diff_2 = self.HasContentsDiff(None,target_sheet_reader,base_sheet_reader)
        if vlu_diff_1 or vlu_diff_2:
            return True
        #内容变更
        vlu_diff_1,dic1,dic2 = self.HasBoundsDiff(None,target_sheet_reader,base_sheet_reader)
        vlu_diff_2,dic1,dic2 = self.HasBoundsDiff(None,target_sheet_reader,base_sheet_reader)
        if vlu_diff_1 or vlu_diff_2:
            return True
        return False
    
    def CompareSheetReader(self,target_reader,base_reader):
        #比较两个Reader的差异，并记录应用的cell的数据和合并单元格的内容
        key_diffs = self._collectKeysDiffType(target_reader,base_reader)
        target_reader.key_diffs = key_diffs
        return key_diffs

    #是否有同一区域的修改
    def HasSameCboundDiff(self,reader1,reader2,interse_set):
        for name in interse_set:
            mine_key_diffs = reader1.sheet_reader_dic[name].key_diffs
            their_key_diffs = reader2.sheet_reader_dic[name].key_diffs
            if mine_key_diffs and their_key_diffs:
                #通过python中的集合来处理 如果没有交集 那就证明没有修改同一处的bounds
                mine_cache_set = set()
                mine_cache_set.update(mine_key_diffs["add_key"])
                mine_cache_set.update(mine_key_diffs["del_key"])
                mine_cache_set.update(mine_key_diffs["mod_key"])
                
                their_cache_set = set()
                their_cache_set.update(their_key_diffs["add_key"])
                their_cache_set.update(their_key_diffs["del_key"])
                their_cache_set.update(their_key_diffs["mod_key"])
                diff_set = mine_cache_set.intersection(their_cache_set)
                if len(diff_set) > 0:
                    mod_cache_dic1 = None
                    mod_cache_dic2 = None
                    if "mod_cache_dic" in mine_key_diffs.keys():
                        mod_cache_dic1 = mine_key_diffs["mod_cache_dic"]
                    if "mod_cache_dic" in their_key_diffs.keys():
                        mod_cache_dic2 = their_key_diffs["mod_cache_dic"]
                    if mod_cache_dic1 is None:
                        mod_cache_dic1 = dict()
                    if mod_cache_dic2 is None:
                        mod_cache_dic2 = dict()
                    return True,name,diff_set,mod_cache_dic1,mod_cache_dic2
        return False,"",diff_set,None,None


    def _collectKeysDiffType(self,target_sheet_reader,base_sheet_reader):
        result = {}
        #方便合并时取cells和合并单元格得信息
        result["target_sheet_reader"] = target_sheet_reader
        result["base_sheet_reader"] = base_sheet_reader
        target_body_list = target_sheet_reader.body_value_list
        base_body_list = base_sheet_reader.body_value_list
        #讲list转换为集合
        target_set = set(target_body_list)
        base_set = set(base_body_list)
        #target对baset的差集 当作新增
        result["add_key"] = target_set.difference(base_set)
        #base对target的差集 当作删除
        result["del_key"] = base_set.difference(target_set)
        #base对target的交集有差异部分当作修改
        interse_set = target_set.intersection(base_set)
        result["mod_key"] = set()
        for key in interse_set:
            diff_state,tar_vlu_dic,bs_vlu_dic = self.HasBoundsDiff(key,target_sheet_reader,base_sheet_reader)
            if diff_state:
                result["mod_key"].add(key)
                if "mod_cache_dic" not in result.keys():
                    result["mod_cache_dic"] = dict()
                if key not in result["mod_cache_dic"].keys():
                    result["mod_cache_dic"][key] = list()
                result["mod_cache_dic"][key].append(tar_vlu_dic)
                result["mod_cache_dic"][key].append(bs_vlu_dic)

        
        result['has_diff'] = False
        if len(result["add_key"]) > 0:
            result['has_diff'] = True
            cache = result["add_key"]
            for item_key in cache:
                DebugHelper.LogColor(DebugHelper.FontColor.green,"【新增】{} {} {}".format(target_sheet_reader.excel_title,target_sheet_reader.sheet.title,item_key))
        
        if len(result["del_key"]) > 0:
            result['has_diff'] = True
            cache = result["del_key"]
            for item_key in cache:
                DebugHelper.LogColor(DebugHelper.FontColor.yellow,"【删除】{} {} {}".format(target_sheet_reader.excel_title,target_sheet_reader.sheet.title,item_key))
        
        if len(result["mod_key"]) > 0:
            result['has_diff'] = True
            cache = result["mod_key"]
            for item_key in cache:
                DebugHelper.LogColor(DebugHelper.FontColor.blue,"【修改】{} {} {}".format(target_sheet_reader.excel_title,target_sheet_reader.sheet.title,item_key))
        return result

    #获取Bounds范围内 行的string信息
    def _getStrValuesDic(self,sheet,bounds):
        cells_info = {}
        index = 0
        for col_items in sheet.iter_rows(bounds.min_row,bounds.max_row,bounds.min_col,bounds.max_col):
            value_str = ""
            for cell in col_items:
                if cell.value != None:                    
                    value_str += str(cell.value)
                    value_str += "___"
                else:
                    value_str += "|NONE|"
            cells_info[index] = value_str
            index = index + 1
        return cells_info

    #判断该bounds是否有差异
    def HasBoundsDiff(self,key,target_sheet_reader,base_sheet_reader):
        target_bounds = None
        base_bounds = None
        if key == None:
            target_bounds = target_sheet_reader.title_bounds
            base_bounds = base_sheet_reader.title_bounds
        else:
            target_bounds = target_sheet_reader.body_bounds_dic[key]
            base_bounds = base_sheet_reader.body_bounds_dic[key]

        target_vluDic = self._getStrValuesDic(target_sheet_reader.sheet,target_bounds)
        base_vluDic = self._getStrValuesDic(base_sheet_reader.sheet,base_bounds)

        for vluStr in target_vluDic.values():            
            if vluStr not in base_vluDic.values():
                return True,target_vluDic,base_vluDic
        for vluStr in base_vluDic.values():            
            if vluStr not in target_vluDic.values():
                return True,target_vluDic,base_vluDic
        return False ,target_vluDic,base_vluDic

        #判断该bounds批注是否有差异
    def HasContentsDiff(self,key,target_sheet_reader,base_sheet_reader):
        if key == None:
            target_bounds = target_sheet_reader.title_bounds
            base_bounds = base_sheet_reader.title_bounds
        else:
            target_bounds = target_sheet_reader.body_bounds_dic[key]
            base_bounds = base_sheet_reader.body_bounds_dic[key]
        str_targrt_content = ""
        str_base_content = ""
        for col_items in target_sheet_reader.sheet.iter_rows(target_bounds.min_row,target_bounds.max_row,target_bounds.min_col,target_bounds.max_col):
            for cell in col_items:
                if cell.comment is not None:
                    str_targrt_content += cell.comment.text
        for col_items in base_sheet_reader.sheet.iter_rows(base_bounds.min_row,base_bounds.max_row,base_bounds.min_col,base_bounds.max_col):
            for cell in col_items:
                if cell.comment is not None:
                    str_base_content += cell.comment.text
        if str_targrt_content != str_base_content:
            return True
        return False

    #判断该bounds合并单元格是否有差异
    def HasMergeCellsDiff(self,key,target_sheet_reader,base_sheet_reader):
        if key == None:
            target_bounds = target_sheet_reader.title_bounds
            base_bounds = base_sheet_reader.title_bounds
        else:
            target_bounds = target_sheet_reader.body_bounds_dic[key]
            base_bounds = base_sheet_reader.body_bounds_dic[key]

        target_mc_list = list()
        base_mc_list = list()
        for merge_cell in target_sheet_reader.sheet.merged_cells.ranges:
            if target_sheet_reader._isInCboundsMergeCell(merge_cell,target_bounds):
                target_mc_list.append(merge_cell.coord)
        for merge_cell in base_sheet_reader.sheet.merged_cells.ranges:
            if base_sheet_reader._isInCboundsMergeCell(merge_cell,base_bounds):
                base_mc_list.append(merge_cell.coord)
        union_mc_list = set(target_mc_list).union(set(base_mc_list))
        if len(union_mc_list) != len(target_mc_list) or len(union_mc_list) != len(base_mc_list):
            
            return True
        return False

    #是否有新增或者删除的sheet
    def GetNewOrDelSheet(self,target_sheet_names,base_sheet_names):
        target_set = set(target_sheet_names)
        base_set = set(base_sheet_names)
        add_set = target_set.difference(base_set)
        del_set = base_set.difference(target_set)
        add_list = list()
        del_list = list()
        if len(add_set) > 0:
            for item in add_set:
                if not IsHasChinese(item):
                    add_list.append(item)
                else:
                    DebugHelper.LogColor(DebugHelper.FontColor.pink,"【新增备注sheet的操作不处理】{} ".format(item))
        if len(del_set) > 0:
            for item in del_list:
                if not IsHasChinese(item):
                    del_list.append(item)
                else:                        
                    DebugHelper.LogColor(DebugHelper.FontColor.pink,"【删除备注sheet的操作不处理】{} ".format(item))
        return len(add_list) > 0 or len(del_list) > 0, add_list,del_list

