from win32com.client import Dispatch
pathExcelFile = r'C:\Users\h215997\OneDrive - Honeywell\send email\templateMassEmail.xlsm'
macroName = 'writeCells'

def callMacro(pathExcelFile,macroName):
    myExcel = Dispatch('Excel.Application')
    myExcel.Visible = 0
    myExcel.Workbooks.Add(pathExcelFile)
    myExcel.Run(macroName)
    myExcel.DisplayAlerts = 0
    myExcel.Quit()

if __name__ == '__main__':
    callMacro(pathExcelFile,macroName)