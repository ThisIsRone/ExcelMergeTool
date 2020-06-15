import pip
from subprocess import call

def loadDown(project_name):
    try:
        call("pip install " + project_name, shell=True)
    except Exception as e:
        print(e)

loadDown("openpyxl")
loadDown("pyinstaller")
