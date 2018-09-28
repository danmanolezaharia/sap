
import pyautogui,time,subprocess
import win32api
import datetime
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

nameLogin = ''
passLogin = ''
serverName ='01.01'
appPath = 'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
waitSecondsToOpen = 5
tCode ='ZFI_DISP_FISCAL_CAL'
fiscalYear = '2018'
waitSecondsToRun = 5


def loginSAP(nameLogin,passLogin):
    pyautogui.typewrite(nameLogin)
    pyautogui.press('tab')
    pyautogui.typewrite(passLogin)
    pyautogui.press('enter')

def selectSapServer(serverName):
    for i in range(0,10):
        pyautogui.press('tab')
    pyautogui.typewrite(serverName)  
    pyautogui.press('enter')


def openApp(appPath,waitSecondsToOpen):
    subprocess.Popen(appPath, stdout=subprocess.PIPE, creationflags=0x08000000)
    time.sleep(waitSecondsToOpen)

def tCodeSap(tCode,fiscalYear,waitSecondsToRun):
    pyautogui.typewrite(tCode)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('tab')
    pyautogui.typewrite(fiscalYear)
    pyautogui.press('f8') #Run
    time.sleep(waitSecondsToRun)

def exportExcel():
    pyautogui.hotkey('alt', 'l')
    pyautogui.press('e')
    pyautogui.press('a')
    time.sleep(2)
    pyautogui.typewrite(tCode + time.strftime("%y%m%d%H%M%S"))
    pyautogui.press('enter')

def exitSAP():
    time.sleep(3)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    pyautogui.hotkey('shift', 'f3')
    time.sleep(2)
    pyautogui.hotkey('shift', 'f3')
    time.sleep(2)
    pyautogui.hotkey('shift', 'f3')
    time.sleep(2)
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')


openApp(appPath,waitSecondsToOpen)
selectSapServer(serverName)
time.sleep(3)
loginSAP(nameLogin,passLogin)
time.sleep(3)
tCodeSap(tCode,fiscalYear,waitSecondsToRun)
exportExcel()
exitSAP()
print("Completed!" + str(i))
time.sleep(10)
