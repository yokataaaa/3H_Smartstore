import win32com.client as win32

def convert_excel(filename):
    fname = r"C:\Users\dlawp\PycharmProjects\project\smartstore\TestTemplate.xls"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    new_fname = r"C:\Users\dlawp\PycharmProjects\project\smartstore\{}".format(filename)
    wb.SaveAs(new_fname, FileFormat = 51) #FileFormat = 51 is for .xlsx extension
    wb.Close() #FileFormat = 56 is for .xls extension
    excel.Application.Quit()





