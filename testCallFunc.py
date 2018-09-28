
import pyautogui,time,subprocess
import win32api
import datetime
from sap import openApp,selectSapServer,loginSAP,tCodeSap,exportExcel,exitSAP
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

openApp(r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe',5)
selectSapServer(r'01.01')
time.sleep(3)
loginSAP(r'gnhbsautomation',r'2%kRvPoep')
time.sleep(3)
tCodeSap(r'ZFI_DISP_FISCAL_CAL',2018,5)
exportExcel()
exitSAP()