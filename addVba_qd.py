# 将bas(vba)文件置入excel文件
# 希望实现选择文件功能
import win32com.client


xls = win32com.client.Dispatch("Excel.Application")

wb = xls.Workbooks.Open('C:\\Users\\shej\
i111\\Desktop\\新桌面\\数据\
分析\\xuzhongting201907121624232486.xls')

wb.VBProject.VBComponents.Import('D:\\cc\
py\\Ex_pyxls\\qudao\\count_qd.bas')

xls.Application.Run('xuzhongting201907121624232486.xls!qdqd1')
wb.Save()
xls.Application.Quit()
