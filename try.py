import win32api, win32gui, win32ui, win32gui_struct, win32con
import unit as ut
import time
import os

filename = '"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\PuTTY (64-bit)"'
os.path.join(filename, 'bin')

os.popen('PuTTY')
time.sleep(1)
windowsName = 'Putty Configuration'
handle = win32gui.FindWindow(None, windowsName)
child = win32gui.FindWindowEx(handle, None, None, 'Host &Name (or IP address)')
ip = win32gui.FindWindowEx(handle, child, 'Edit', None)

# mouse event
rect = win32gui.GetWindowRect(ip)
x, y = rect[0], rect[1]
win32api.SetCursorPos([x, y])
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

# key event
win32api.keybd_event(ut.VK_CODE['ctrl'], 0, 0, 0)
win32api.keybd_event(ut.VK_CODE['a'], 0, 0, 0)
win32api.keybd_event(ut.VK_CODE['a'], 0, win32con.KEYEVENTF_KEYUP, 0)
win32api.keybd_event(ut.VK_CODE['ctrl'], 0, win32con.KEYEVENTF_KEYUP, 0)
win32api.keybd_event(ut.VK_CODE['del'], 0, 0, 0)
win32api.keybd_event(ut.VK_CODE['del'], 0, win32con.KEYEVENTF_KEYUP, 0)

time.sleep(1)
string = b''
for x in string: #依次发送字节串中的每个字节
	win32gui.SendMessage(ip, win32con.WM_CHAR, x, 0)


open_b = win32gui.FindWindowEx(handle, None, None, '&Open')
rect = win32gui.GetWindowRect(open_b)
x, y = rect[0], rect[1]
win32api.SetCursorPos([x, y])
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
time.sleep(0.5)
pu_str = ' - PuTTY'
putty = win32gui.FindWindow(None, pu_str)
print(putty)
ut.key_bd('shift')
ut.key_bd('o')
ut.key_bd('n')
ut.key_bd('o')
ut.key_bd('s')
ut.key_bd('enter')
time.sleep(0.5)
ut.key_bd('i')

ut.key_bd('enter')
