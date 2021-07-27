import win32gui
import win32ui
from ctypes import windll
from PIL import Image
import re
import time
import cv2
import pytesseract
import pywinauto
import numpy as np
from pynput.keyboard import Key, Controller
from random import randint as rand
import win32com.client
import os



class WindowMgr:
    def __init__(self):
        self._handle = None

    def find_window(self, class_name, window_name=None):
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
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


pytesseract.pytesseract.tesseract_cmd = r'D:\\Programs\\Tesseract\\tesseract.exe'
windows = pywinauto.Desktop(backend="uia").windows()
discord_title = [w.window_text()
                 for w in windows if "Discord" in w.window_text()][0]
hwnd = win32gui.FindWindow(None, discord_title)
w = WindowMgr()
w.find_window_wildcard(discord_title)
keyboard = Controller()
shell = win32com.client.Dispatch("WScript.Shell")


lower = np.array([200, 200, 200], dtype="uint16")
upper = np.array([255, 255, 255], dtype="uint16")


def get_text(img):
    img_array = np.asarray(img)
    cropped_img = img_array[200:990, 450:620]
    scaled_img = scale_image(cropped_img, 1000)
    black_white_img = cv2.inRange(scaled_img, lower, upper)
    eroded_img = cv2.erode(black_white_img, np.ones((4, 4), np.uint8))
    filtered_img = cv2.filter2D(eroded_img, -1, light_blur_kernel)
    text = pytesseract.image_to_string(filtered_img)
    return text, filtered_img
    

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
        l[i] = l[i].replace("?", "7")
        l[i] = l[i].replace("l", "I")
        
        #TODO: this can be removed if it creates problems:
        # l[i] = l[i].replace("O", "0") 
        # l[i] = l[i].replace("0", "O")
    return l


def set_focus(hwnd):
    shell.SendKeys('%') 
    win32gui.SetForegroundWindow(hwnd)
    



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

for i in range(20000):
    img = save_image(hwnd)
    text = pytesseract.image_to_string(img)
    if "to claim it!" in text:
        count = 0
        while True:
            print("MATCH FOUND")
            text, final_img = get_text(img)
            matches = re.findall(r"\b\w{4}\b", text)
            matches = replace_characters(matches)

            if len(matches) > 0:
                set_focus(hwnd)
                code = matches[-1].lower()
                type_string(keyboard, "!claim " + code, True)
                print(matches)
                cv2.imwrite(code + ".png" , final_img)
                break

            else:
                print("NO CODES FOUND")
                cv2.imwrite("broken" + str(rand(1, 10000)) + ".png", np.asarray(final_img))
                img = save_image(hwnd)
                count += 1
                time.sleep(0.02)
                if count > 3:
                    break
                print(text)
        print("--------------------")
        time.sleep(160)

    time.sleep(0.05)

    # w.set_foreground()
# TODO: refactor
# TODO: have successful/unsuccessful go in different folders with transcript etc
# TODO: remove spaces
# TODO: add automatic view switching


# Problems:
# Too slow
# Sometimes doesn't recognise the codes at all
# 0s and Os, 5s and 9s mixed up (maybe more)