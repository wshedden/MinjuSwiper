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

    def __init__(self):
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

    def set_focus(self):
        win32gui.SetFocus(self._handle)


def save_image(hwnd):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
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


blur_kernel = np.ones((10, 10), np.float32)/100
light_blur_kernel = np.ones((5, 5), np.float32)/25
erode_kernel = np.ones((3, 3), np.uint8)


def scale_image(img, scale_percent):
    width = int(img.shape[1] * scale_percent / 100)
    height = int(img.shape[0] * scale_percent / 100)
    dim = (width, height)
    img = cv2.resize(img, dim)
    img = cv2.filter2D(img, -1, blur_kernel)
    return img

def replace_characters(l):
    for i in range(len(l)):
        l[i].replace("?", "7")
        l[i].replace("l", "I")

    return l


def set_focus(w):
    # w.set_foreground()
    pass


pytesseract.pytesseract.tesseract_cmd = r'D:\\Programs\\Tesseract\\tesseract.exe'
windows = Desktop(backend="uia").windows()
discord_title = [w.window_text()
                 for w in windows if "Discord" in w.window_text()][0]
hwnd = win32gui.FindWindow(None, discord_title)
w = WindowMgr()
w.find_window_wildcard(discord_title)
keyboard = Controller()


lower = np.array([200, 200, 200], dtype="uint16")
upper = np.array([255, 255, 255], dtype="uint16")



def test():
    img = cv2.imread("ex1.png")
    img2 = cv2.imread("ex2.png")
    
    # erode
    
    img = cv2.erode(img, np.ones((4, 4), np.uint8))
    img = cv2.filter2D(img, -1, light_blur_kernel)

    img2 = cv2.erode(img2, np.ones((4, 4), np.uint8))
    img2 = cv2.filter2D(img2, -1, light_blur_kernel)

    text = pytesseract.image_to_string(img)
    text2 = pytesseract.image_to_string(img2)
    print(text, text2)

    cv2.imshow("test", img)
    cv2.imshow("test", img2)

    cv2.waitKey()
    cv2.destroyAllWindows()



test()
input()



# set_focus(w)
# time.sleep(2)
# set_focus(w)

# print("done")
# input()

for i in range(20000):
    img = save_image(hwnd)
    text = pytesseract.image_to_string(img)
    if "to claim it!" in text:
        count = 0
        while True:
            print("MATCH FOUND")

            # cv2.imwrite("example_raw.png", np.asarray(img))
            # time.sleep(0.1)
            # cv2.imwrite("example_raw2.png", np.asarray(img))
            # time.sleep(0.1)
            # cv2.imwrite("example_raw3.png", np.asarray(img))
            # time.sleep(0.1)

            # exit

            img_array = np.asarray(img)
            # img_rgb = cv2.cvtColor(img_array, cv2.COLOR_BGR2RGB)
            # not using RGB
            crop_img = img_array[700:990, 450:620]

            #scale, erode twice and blur
            scaled_image = scale_image(crop_img, 500)
            
            # scaled_image = scale_image(scaled_image, 200)

            # scaled_image = cv2.erode(scaled_image, np.ones((3, 3), np.uint8))
            # scaled_image = cv2.filter2D(scaled_image, -1, light_blur_kernel)

            

            filtered_image = cv2.inRange(scaled_image, lower, upper)
            eroded_image = cv2.erode(filtered_image, erode_kernel)


            text2 = pytesseract.image_to_string(filtered_image)

            matches = re.findall(r"\b\w{4}\b", text2)
            matches = replace_characters(matches)

            if len(matches) > 0:
                set_focus(w)
                code = matches[-1].lower()
                type_string(keyboard, "!claim " + code, True)
                print(matches)
                print(matches[-1])
                cv2.imwrite(code + ".png", filtered_image)
                break

            else:
                print("NO CODES FOUND")
                n = str(random.randint(1, 10000))
                cv2.imwrite("broken" + n + str(count) + 
                            ".png", filtered_image)
                cv2.imwrite("broken_scaled" + n + str(count) + 
                            ".png", scaled_image)

                text = pytesseract.image_to_string(crop_img)
                matches = re.findall(r"\b\w{4}\b", text)
                img = save_image(hwnd)
                text = pytesseract.image_to_string(img)

                count += 1
                time.sleep(0.05)
                if count > 1:
                    break
                print(text2)
        print("--------------------")
        time.sleep(150)

    time.sleep(0.1)

    # w.set_foreground()
# TODO: refactor
# TODO: have successful/unsuccessful go in different folders with transcript etc
# TODO: remove spaces
# TODO: add automatic view switching
