 """
 * @author Rone Cao
 *
 * @email 13592468626@163.com
 *
 * @create date 2020-02-28 19:19:23
 """
import sys
from CopyHelper import FileCopyHelper
from ExcelMergeMain import ExcelMergeMain
import DebugHelper
# D:\Users\admin\Anaconda3\python.exe "D:\SvnProject\Tools\ExcelMergeTool\Diff3Warp.py" D:\SvnProject\Excel_Merge_Tool_Test\Copy_test\TestMachine.xlsm D:\SvnProject\Excel_Merge_Tool_Test\Copy_test\TestMachine.xlsm.r3 D:\SvnProject\Excel_Merge_Tool_Test\Copy_test\TestMachine.xlsm D:\SvnProject/Excel_Merge_Tool_Test\Copy_test\TestMachine.xlsm.r1
if __name__ == '__main__':
    path_modify = dict(
        base = sys.argv[-1],
        mine = sys.argv[-2],
        their = sys.argv[-3],
        merge = sys.argv[-4],
    )
    mine_copy_helper = FileCopyHelper(path_modify["base"])
    their_copy_helper = FileCopyHelper(path_modify["their"])
    path_modify["base"] = mine_copy_helper.copypath
    path_modify["their"] = their_copy_helper.copypath

    merger = ExcelMergeMain(path_modify["merge"],path_modify["their"],path_modify["mine"],path_modify["base"])
    result = merger.StartWork()
    merger.OnRelease()
    if result:
        #如果检查通过 并合并 那就把base和their删除掉
        DebugHelper.Log("【合并成功】")
        mine_copy_helper.DelSourceFile()
        their_copy_helper.DelSourceFile()
    mine_copy_helper.OnRelease()
    their_copy_helper.OnRelease()
    DebugHelper.Log("输入exit关闭当前窗口")
    active = True
    while active:
        message = input()
        if message == "exit":
            active = False