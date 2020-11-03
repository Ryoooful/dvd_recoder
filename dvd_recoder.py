import pyautogui
import time
import win32gui
import win32con
import subprocess

windowclass = 'PowerDVD14'
windowname = None

h = 0
m = 51
s = 0

m += h * 60
s += m * 60


def volume_control():
  subprocess.Popen('sndvol.exe')
  time.sleep(1)
  volumeclass = '#32770'
  winname = None
  volumeWindows = win32gui.FindWindow(volumeclass,winname)
  if volumeWindows is not 0:
    win32gui.SetForegroundWindow(volumeWindows)
    time.sleep(1)
    pyautogui.hotkey('home', 'pagedown')
    time.sleep(1)
    pyautogui.hotkey('pagedown')

def recode_start():
  hWnd = win32gui.FindWindow(windowclass,windowname)
  if hWnd is not 0:
    x, y, xx, yy =win32gui.GetWindowRect(hWnd)
    pyautogui.click(x + 10, y + 10)
    win32gui.SetForegroundWindow(hWnd)
    pyautogui.hotkey('winleft','g')
    time.sleep(1)
    pyautogui.hotkey('winleft','alt','r')
    pyautogui.click(xx - 10, yy - 10)

def recode_stop():
  hWnd = win32gui.FindWindow(windowclass,windowname)
  if hWnd is not 0:
    x, y, xx, yy =win32gui.GetWindowRect(hWnd)
    pyautogui.click(x + 10, y + 10)
    win32gui.SetForegroundWindow(hWnd)
    pyautogui.hotkey('winleft','alt','r')
    win32gui.PostMessage(hWnd, win32con.WM_CLOSE,0,0)



volume_control()
recode_start()
time.sleep(s)
recode_stop()
