import pyautogui
import time,sys
import win32gui
#import pywin32
import os
#### Prerequisite:
####    python -m pip install PyAutoGUI
####    python -m pip install Pillow
####    python -m pip install pypiwin32, win32gui

start = 2
end = 122
dir_book = "C:\\Users\\itqiuw\\Desktop\\D75AW"         # notice "\\"
#dir_book = "C:\\Users\\Administrator\\Desktop\\tadmd1"

pyautogui.PAUSE = 1           # delay x seconds for all the functions

screenWidth, screenHeight = pyautogui.size()
pyautogui.moveTo(screenWidth / 2, screenHeight / 2)

print("Screen size: " + str(screenWidth) + " x " + str(screenHeight))

#png_folder = "1366X768\\"
png_folder = str(screenWidth) + "X" + str(screenHeight) + "\\"
print_button = png_folder + "print_button.png"
nextpage = png_folder + "nextpage_button.png"
save = png_folder + "save_button_zh.png"
page_input = png_folder + "page_input.png"


def ClickPng(png,qty=1):
    tcount = 0
    while tcount < 500:
        try:
            x, y = pyautogui.locateCenterOnScreen(png)
            if qty == 2:
                pyautogui.doubleClick(x, y)
            elif qty == 1:
                pyautogui.click(x, y)
            break
        except:
            time.sleep(0.1)
            tcount = tcount + 1
            print("png scanning...")
    if tcount >= 500:
        print("button " + png + " not found")
        sys.exit()

def Find_Window(name):
    hwnd = win32gui.FindWindow(None,name)
    tmptime = 0
    if hwnd:
        win32gui.SetForegroundWindow(hwnd)
        return 1
    else:
        return 0

def Poll_page(page):
    ClickPng(print_button)        # Click Print Button
    ClickPng(save)
    tmptime = 0
    while tmptime < 50:         # timeout 50*0.5 seconds
        if Find_Window('另存为'):
            break
        if Find_Window('Save as'):
            break
        tmptime = tmptime + 1
        if tmptime == 50:
            pyautogui.hotkey('alt', 'f4')
            relist.append(page)
            print(str(page) + "not respond.")
            return
        time.sleep(0.5)
        
    pyautogui.typewrite(str(page), interval=0.1)
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'f4')


####################################################
############# Main Program #########################

relist = []

for i in range(start,end+1,1):
    ClickPng(nextpage)
    Poll_page(i)

## Re-print the blank pages

for parent,dirnames,filenames in os.walk(dir_book):
    for filename in filenames:
        tmppage = int(filename[:-4])
        fullfile = dir_book + "\\" + filename
        Size=os.path.getsize(fullfile)
        if Size < 1024 and tmppage > 2:
            relist.append(tmppage)
            print("Size of " + filename + ":\t" + str(Size))
            os.remove(fullfile)

print(relist)
print("Totally " + str(len(relist)) + " slides to re-poll...")


while len(relist):
    tmppage = relist.pop()
    print(tmppage)
    ClickPng(page_input,2)         # why this png can't be identified?
##    pyautogui.doubleClick(1740, 1055)
##    pyautogui.doubleClick(1204, 746)
    pyautogui.typewrite(str(tmppage), interval=0.1)
    pyautogui.press('enter')
    Poll_page(tmppage)

print("Polling Completed!")
