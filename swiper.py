import win32gui
import win32ui
from ctypes import windll
from PIL import Image
import re
import time
import cv2
import pytesseract
from pywinauto import Desktop
import numpy as np
from pynput.keyboard import Key, Controller
import random

class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)


def save_image(hwnd):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)

    saveDC.SelectObject(saveBitMap)

    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    if result == 1:
        return im

def type_string(kb, s, press_enter):
    for i in s:
        kb.press(i)
        kb.release(i)
    if press_enter:
        kb.press(Key.enter)
        kb.release(Key.enter)



def filter_image(img):
    return img


pytesseract.pytesseract.tesseract_cmd = r'D:\\Programs\\Tesseract\\tesseract.exe'



img = cv2.imread("broken9466.png")
kernel = np.ones((2, 2),np.float32)/4
dst = cv2.filter2D(img,-1,kernel)
cv2.imshow("test", dst)
cv2.waitKey()
cv2.destroyAllWindows()
text = pytesseract.image_to_string(img)
print(text)

input()




windows = Desktop(backend="uia").windows()
discord_title = [w.window_text() for w in windows if "Discord" in w.window_text()][0]
hwnd = win32gui.FindWindow(None, discord_title)
w = WindowMgr()
w.find_window_wildcard(discord_title)
keyboard = Controller()


for i in range(20000):
    img = save_image(hwnd)
    # img = cv2.imread("test.png")
    text = pytesseract.image_to_string(img)
    if "to claim it!" in text:
        print("MATCH FOUND")
        # Filter image
        # img = filter_image(img)
        img_array = np.asarray(img)
        img_rgb = cv2.cvtColor(img_array, cv2.COLOR_BGR2RGB)
        # crop_img = img_rgb[100:990, 350:720]
        crop_img = img_rgb[850:990, 350:720]

        lower = np.array([180, 180, 180], dtype = "uint16")
        upper = np.array([255, 255, 255], dtype = "uint16")

        filtered_image = cv2.inRange(crop_img, lower, upper)

        # cv2.imshow("IMAGE", filtered_image)
        # cv2.waitKey()
        # cv2.destroyAllWindows()
        # text = pytesseract.image_to_string(crop_img)
        # print(text)
        # input()
        text2 = pytesseract.image_to_string(filtered_image)

        matches = re.findall(r"\b\w{4}\b", text2)
        
        if len(matches) > 0:
            print(matches)
            code = matches[-1].lower()
            print(code)
            type_string(keyboard, "!claim " + code, True)
            cv2.imwrite(code + ".png", filtered_image)

        else:
            print("NO CODES FOUND")
            cv2.imwrite("broken" + str(random.randint(1, 10000)) + ".png", filtered_image)

            text = pytesseract.image_to_string(crop_img)
            matches = re.findall(r"\b\w{4}\b", text)

        
        print(text2)

        # cv2.imshow("IMAGE", filtered_image)
        # cv2.waitKey()
        # cv2.destroyAllWindows()

        time.sleep(60)

    time.sleep(0.1)

    # w.set_foreground()
