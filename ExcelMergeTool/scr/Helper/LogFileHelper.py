import os

#log的日志文件
log_file  = None
log_file_name = None
prefix = "合并工具日志文件.txt"

def IsExist(mine_file_path):
    log_file_name = os.path.basename(mine_file_path)
    log_file_name = log_file_name + prefix
    return os.path.exists(log_file_name)

def InitLogFile(mine_file_path):
    global log_file
    global log_file_name
    log_file_name = os.path.basename(mine_file_path)
    log_file_name = log_file_name + prefix
    log_file = open(log_file_name,'w')

def ReleaseLogFile():
    global log_file
    global log_file_name
    log_file.close()
    log_file = None

def DelLogFile():
    if os.path.exists(log_file_name):
        os.remove(log_file_name)