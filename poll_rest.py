import pyautogui
import time,sys

start = 19
end = 138
relist = []

pyautogui.PAUSE = 1           # delay x seconds for all the functions
##relist = [18,34,50,66,82,98,114,130,146,162,178,194,210,226]

print_button = "print_button.png"
nextpage = "nextpage_button.png"
save = "save_button_zh.png"


import os
import os.path
import sys

start = 2
end = 138
dir_book = "D:\\adm800"         # notice "\\"
relist = []


##for i in range(2,end,2):
##    relist.append(i)
##
##for parent,dirnames,filenames in os.walk(dir_book):
##    for filename in filenames:
##        tmp = int(filename[:-4])
##        if tmp > 1 and tmp < end-1:
##            relist.remove(tmp)
##    print("Totally" + str(len(relist)) + " slides to poll...")

##sys.exit()



#time.sleep(0.5)

screenWidth, screenHeight = pyautogui.size()
pyautogui.moveTo(screenWidth / 2, screenHeight / 2)


def ClickPng(png):
    tcount = 0
    while tcount < 50:
        try:
            x, y = pyautogui.locateCenterOnScreen(png)
            pyautogui.click(x, y)
            break
        except:
            time.sleep(0.1)
            tcount = tcount + 1
    if tcount >= 50:
        print("button " + png + "not found")
        sys.exit()
            
    



def Poll_page(page, double = True):

    ClickPng(nextpage)
    ClickPng(print_button)        # Click Print Button
    ClickPng(save)

    #time.sleep(0.5)
    pyautogui.typewrite(str(page), interval=0.1)
    pyautogui.press('enter')

    #avoid replace action
    #pyautogui.press('left')
    #pyautogui.press('enter')

    pyautogui.hotkey('alt', 'f4')

    
    #time.sleep(0.5)



############# Main Program ########################

for i in range(start,end+1,1):
    print(i)
    Poll_page(i)


#pyautogui.click(clicks=2, interval=0.25)


print("Polling Completed!")
