# 连接到目前活跃的文档
import win32com.client as win32
win_cad = win32.Dispatch("AutoCAD.application")
doc = win_cad.ActiveDocument
msp = doc.ModelSpace

# 在活跃的文档的公屏上打上hello world
doc.Utility.prompt("Hello world!\n")
# 在pycharm中打印文档名称
print(doc.name)
