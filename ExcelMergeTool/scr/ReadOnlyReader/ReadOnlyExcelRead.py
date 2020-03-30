# -*- coding: utf-8 -*-
"""
Created on Fri Dec 13 15:12:40 2019

@author: admin
"""

from openpyxl import load_workbook
from ReadOnlySheetReader import ReadOnlySheetReader as SheetReader
import scr.Helper.DebugHelper as DebugHelper

class ReadOnlyExcelRead:
        
    def __init__(self,path_xlsm,excel_title):
        self.excel_title = excel_title
        self.path_xlsm = path_xlsm
        if path_xlsm is None:
             raise Exception("Invalid ExcelReader path_xlsm!", path_xlsm)
        DebugHelper.Log("【读取资源】 " + path_xlsm + " " + excel_title)
        self.workBook = load_workbook(filename = path_xlsm, read_only=False, keep_vba=True,data_only=True)
        sheet_names = self.workBook.get_sheet_names()
        sheet_reader_dic = dict()
        for name in sheet_names:
            sheet = self.workBook.get_sheet_by_name(name)
            reader = SheetReader(sheet,excel_title)
            sheet_reader_dic[name] =  reader
        self.sheet_reader_dic = sheet_reader_dic
        self.sheet_names = set(sheet_names)
        self.has_diff = False
        DebugHelper.Log("【读取完成】 " + path_xlsm)

    def OnRelease(self):
        self.workBook.close()
        self.workBook = None
        self.path_xlsm = None
        self.sheet_names = None
        self.sheet_reader_dic = None
        