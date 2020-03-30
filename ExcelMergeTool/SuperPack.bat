@echo off
set "dist_dir=%cd%\dist"
set "build_dir=%cd%\build"
set "pycache_dir=%cd%\__pycache__"
set "spec_file=%cd%\ExcelMerge.spec"

echo del dir %dist_dir%
rd /s /q %dist_dir%
::打包exe命令
pyinstaller -F MergeToolApp.py -n ExcelMerge
::删除build文件夹
echo del dir %build_dir%
rd /s /q %build_dir%
::删除__pycache__文件夹
echo del dir %pycache_dir%
rd /s /q %pycache_dir%
::删除ExcelMerge.spec文件
echo del file %spec_file%
del %spec_file%
pause