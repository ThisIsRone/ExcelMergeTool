# -*- coding: utf-8 -*-
"""
Created on Fri Dec 13 15:03:01 2019

@author: admin

ExcelMergeTool
ExcelReader
SheetReader
CompareBounds:记录sheet的数据对应区域的有效范围
"""
from .ExcelReader import ExcelReader
from .CompareSheetReader import CompareSheetReader
from scr.Helper.SheetCopy import CopySheet
import scr.Helper.DebugHelper as DebugHelper

class ExcelMergeMain:
    
    def __init__(self,path_merge,path_their,path_mine,path_base):
        
        self.path_merge = path_merge
        self.path_their = path_their
        self.path_mine = path_mine
        self.path_base = path_base
        
        self.merge_reader = ExcelReader(path_merge,"Merge")
        self.their_reader = ExcelReader(path_their,"Their")
        self.mine_reader = ExcelReader(path_mine,"Mine")
        self.base_reader = ExcelReader(path_base,"Base")
        self.comparer = CompareSheetReader()
    
    def StartWork(self):
        # mine=> base的比较
        DebugHelper.Log("【收集差异】Mine 与 Base")
        self.CheckAndUpdateExcelDiff(self.mine_reader,self.base_reader)
        # their=> base的比较
        DebugHelper.Log("【收集差异】Their 与 Base")
        self.CheckAndUpdateExcelDiff(self.their_reader,self.base_reader)
        state = self.IsSupportCurrentMerge()
        if state:
            self.ApplyDiff2MergeExcel()
            self.CheckAndApplySheetOp()
        else:
            DebugHelper.Log("【合并失败】不支持的类型")
        return state

    #获取两Excel中具有相同名字的意思
    def _getSameSheetNames(self,target_reader,base_reader):
        target_names = target_reader.sheet_names
        base_names = base_reader.sheet_names
        interse_set = target_names.intersection(base_names)
        return interse_set


    #检查并更新Excel的差异信息
    def CheckAndUpdateExcelDiff(self,target_reader,base_reader):
        #TODO 新增sheet
        sheet_names = self._getSameSheetNames(target_reader,base_reader)
        for name in sheet_names:
            mine_sheet = target_reader.sheet_reader_dic[name]
            base_sheet = base_reader.sheet_reader_dic[name]
            if mine_sheet == None or base_sheet == None:
                continue
            self.comparer.CompareSheetReader(mine_sheet,base_sheet)
            # result = self.comparer.CompareSheetReader(mine_sheet,base_sheet)
            # if not target_reader.has_diff:
            #     target_reader.has_diff = result["has_diff"]
            

    #是否支持当前的合并
    def IsSupportCurrentMerge(self):
        mine_name_set = set(self.mine_reader.sheet_names)
        their_name_set = set(self.their_reader.sheet_names)
        interse_set = mine_name_set.intersection(their_name_set)
        #不支持表头有变更
        for name in interse_set:

            if name in self.mine_reader.sheet_reader_dic.keys() and name in self.base_reader.sheet_reader_dic.keys():
                sheet_reader1 = self.mine_reader.sheet_reader_dic[name]
                sheet_reader2 = self.base_reader.sheet_reader_dic[name]
                if self.comparer.IsHasTitleDiff(sheet_reader1,sheet_reader2):
                    DebugHelper.LogColor(DebugHelper.FontColor.red,"【表头变更】终止合并 Mine和Base {}比较 存在表头差异(批注，合并单元格状态，占用行列，数据) ".format(name))
                    sheet_reader1.PrintTitle()
                    sheet_reader2.PrintTitle()
                    return False

            if name in self.their_reader.sheet_reader_dic.keys() and name in self.base_reader.sheet_reader_dic.keys():
                sheet_reader1 = self.their_reader.sheet_reader_dic[name]
                sheet_reader2 = self.base_reader.sheet_reader_dic[name]
                if self.comparer.IsHasTitleDiff(sheet_reader1,sheet_reader2):
                    DebugHelper.LogColor(DebugHelper.FontColor.red,"【表头变更】终止合并 Their和Base {}比较 存在表头差异（(批注，合并单元格状态，占用行列，数据）".format(name))
                    sheet_reader1.PrintTitle()
                    sheet_reader2.PrintTitle()
                    return False

            if name in self.their_reader.sheet_reader_dic.keys() and name in self.mine_reader.sheet_reader_dic.keys():
                sheet_reader1 = self.their_reader.sheet_reader_dic[name]
                sheet_reader2 = self.mine_reader.sheet_reader_dic[name]
                if self.comparer.IsHasTitleDiff(sheet_reader1,sheet_reader2):
                    DebugHelper.LogColor(DebugHelper.FontColor.red,"【表头变更】终止合并 Their和Mine {}比较 存在表头差异（(批注，合并单元格状态，占用行列，数据）".format(name))
                    sheet_reader1.PrintTitle()
                    sheet_reader2.PrintTitle()
                    return False
        #合并操作需要只处理两个目标版本的差异数据，暂不判断同一修改区域的逻辑
        #不支持 修改mine和thier修改了同一张sheet的同一key的bounds下的数据
        interse_set = self._getSameSheetNames(self.mine_reader,self.their_reader)
        has_same,sheet_name,diff_keys,mine_diff_dic,their_diff_dic = self.comparer.HasSameCboundDiff(self.mine_reader,self.their_reader,interse_set)
        if has_same:
            for same_key in diff_keys:
                DebugHelper.LogColor(DebugHelper.FontColor.red,"【重复变更】Thier和Mine出现了同一区域的修改 {} {}".format(sheet_name,same_key))
                self.PrintDiffDic(mine_diff_dic,their_diff_dic,same_key)
            return False
        return True

    def ApplyDiff2MergeExcel(self):
        mine_reader = self.mine_reader
        their_reader = self.their_reader
        sheet_names = self._getSameSheetNames(mine_reader,their_reader)
        self.merge_reader.ApplyDiff2MergeSheet(mine_reader,their_reader,sheet_names)
    
    def PrintDiffDic(self,mine_diff_dic,their_diff_dic,same_key):
        if same_key in mine_diff_dic.keys():
            array1 = mine_diff_dic[same_key]
            mine_diffs = array1[0]
            base_diffs = array1[1]
            for diff in mine_diffs.values():
                DebugHelper.Log(" 【mine】",same_key,":",end = "")
                self._printStrValue(diff)
                DebugHelper.LogNoTime("")
            for diff in base_diffs.values():
                DebugHelper.Log(" 【Base】",same_key,":",end = "")
                self._printStrValue(diff)
                DebugHelper.LogNoTime("")
        if same_key in their_diff_dic.keys():
            array1 = their_diff_dic[same_key]
            their_diffs = array1[0]
            for diff in their_diffs.values():
                DebugHelper.Log("【Their】",same_key,":",end = "")
                self._printStrValue(diff)
                DebugHelper.LogNoTime("")

    def _printStrValue(self,str_vlu):
        str_array = str_vlu.split("___")
        for vlu in str_array:
            DebugHelper.LogNoTime(vlu,"\t",end = "")

    def CheckAndApplySheetOp(self):
        merge_reader = self.merge_reader
        their_reader = self.their_reader
        #获取并处理新增或删除的sheet
        has_diff,add_list,del_list = self.comparer.GetNewOrDelSheet(their_reader.sheet_names,self.merge_reader.sheet_names)
        if has_diff:
            if len(add_list) > 0:
                DebugHelper.Log("【开始新增sheet】")
                for sheet_name in add_list:
                    merge_reader.has_diff = True
                    if sheet_name not in merge_reader.sheet_names:
                        DebugHelper.Log("【新增sheet】Sheet:{}".format(sheet_name))
                        target_sheet = merge_reader.workBook.create_sheet(sheet_name)
                        source_sheet = their_reader.sheet_reader_dic[sheet_name]
                        CopySheet(target_sheet,source_sheet.sheet)
                        DebugHelper.Log("【新增sheet】Sheet:{} Done".format(sheet_name))
                    else:
                        DebugHelper.LogColor(DebugHelper.FontColor.red,"【新增sheet失败】Their和Mine都新增了{} Sheet ".format(sheet_name))
            if len(del_list) > 0:
                DebugHelper.Log("【开始删除sheet】")
                for sheet_name in del_list:
                    merge_reader.has_diff = True
                    if sheet_name in merge_reader.sheet_names:
                        DebugHelper.LogColor(DebugHelper.FontColor.yellow,"【不支持删除sheet】Sheet: {}".format(sheet_name))
                    #     merge_reader.workBook.remove(sheet_name) 
                    #     DebugHelper.Log("【删除sheet】Sheet: {} Done".format(sheet_name))
                    # else:
                    #     DebugHelper.LogColor(DebugHelper.FontColor.red,"【删除sheet失败】Their和Mine都删除了{} Sheet ".format(sheet_name))


    #合并完成
    def OnRelease(self):
        self.merge_reader.OnRelease()
        self.their_reader.OnRelease()
        self.mine_reader.OnRelease()
        self.base_reader.OnRelease()