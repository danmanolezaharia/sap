
import pyautogui,time,subprocess
import win32api
import datetime
from sap import openApp,selectSapServer,loginSAP,tCodeSap,exportExcel,exitSAP
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

nameLogin = 'H215997'
passLogin = 'Parola1234r321'
serverName ='01.01'
appPath = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
waitSecondsToOpen = 5
tCode ='ZFI_DISP_FISCAL_CAL'
fiscalYear = '2020'
waitSecondsToRun = 5

openApp(appPath,waitSecondsToOpen)
selectSapServer(serverName)
time.sleep(3)
loginSAP(nameLogin,passLogin)
time.sleep(3)
tCodeSap(tCode,fiscalYear,waitSecondsToRun)
exportExcel()
exitSAP()