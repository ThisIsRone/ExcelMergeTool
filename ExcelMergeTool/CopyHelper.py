"""
* @author Rone Cao
*
* @email 13592468626@163.com
*
* @create date 2020-02-28 19:19:23
在SVN的冲突的情况下，冲突的后缀会设置为.r1 .r2
这种情况下 openpyxl无法识别后缀为 .r1 .r2的文件
故采用复制修改名字的办法来解决这个错误
"""

from shutil import copyfile
import os, sys

instance_id = 0
class FileCopyHelper:
      def __init__(self,source_path):
         global instance_id
         instance_id += 1

         self.suffix = "file_copy_here_{}.xlsm".format(instance_id)
         self.copy_path = None
         self.source_path = source_path
         self._execute_copy()

      #复制文本并修改名字
      def _execute_copy(self):
         source_path = self.source_path
         if source_path.endswith(".r3"):
            self.copy_path = source_path.replace(".xlsm.r3",self.suffix)
         if source_path.endswith(".r1"):
            self.copy_path = source_path.replace(".xlsm.r1",self.suffix)
         if self.copy_path != None:
            copyfile(self.source_path, self.copy_path)

      @property
      def copypath(self):
         if not self.copy_path:
            return ""
         else:
            return self.copy_path

      def DelSourceFile(self):
         if self.source_path:
            #判断文件是否存在
            if(os.path.exists(self.source_path)):
               os.remove(self.source_path)

      def OnRelease(self):
         if self.copy_path:
            #判断文件是否存在
            if(os.path.exists(self.copy_path)):
               os.remove(self.copy_path)

# file_copy = FileCopyHelper("D:/SvnProject/Excel_Merge_Tool_Test/Copy/TestMachine.xlsm.r3")
# active = True
# while active:
#     message = input()
#     if message == "continue":
#         active = False
# file_copy.OnRelease()