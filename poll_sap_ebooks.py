import pyautogui
import time,sys,os
import win32gui
import configparser
from pathlib import Path
from PyPDF2 import PdfFileReader, PdfFileWriter
#### Prerequisite:
#### (--index-url https://pypi.douban.com/simple)
####    python -m pip install pyautogui
####    python -m pip install Pillow
####    python -m pip install pywin32 [win32gui is a part of it]
####    python -m pip install pypdf2
#### v3.0   add screen size detection mechanism
#### v4.0   add pdf merge function
#### v4.1   add re-poll mode
#### v5.0   add interaction input

version = "5.0"

start = 2
end = 0
dir_book = ""
pa_name = ""
is_repoll = ""
book_path = "C:\\Users\\" + os.environ['USERNAME'] + "\\Desktop\\"
config_file = 'target_book.cfg'

#Initial Variables
pyautogui.PAUSE = 1           # delay x seconds for all the functions
pyautogui.FAILSAFE = False

screenWidth, screenHeight = pyautogui.size()
pyautogui.moveTo(screenWidth / 2, screenHeight / 2)

print("Screen size: " + str(screenWidth) + " x " + str(screenHeight))
png_folder = str(screenWidth) + "X" + str(screenHeight) + "\\"           #choose the screen pixels folder
print_button = png_folder + "print_button.png"
nextpage = png_folder + "nextpage_button.png"
save = png_folder + "save_button_zh.png"
page_input = png_folder + "page_input.png"

def Get_input():
    global dir_book,pa_name,end,is_repoll
    is_repoll = "No"
    cnf = configparser.ConfigParser()
    tmp_file = Path(config_file)
    if tmp_file.is_file():
        print("Config file detected, re-poll mode will be enabled...")
        is_repoll = "Yes"

        cnf.read(config_file)
        dir_book = cnf.get('book', 'bookfolder')           #get(section, option)
        pa_name = cnf.get('book', 'filename')              
        end = cnf.get('book', 'pages')                     #page end
    else:
        print("It's a new polling...")
        dir_book = input("Please input the directory which all seperated pdfs saved from Chrome and it's on desktop:\n")
        pa_name = input("Please input the book full name which will be used as final pdf name:\n")
        end = input("Please input the end page of the book:\n")

        if not cnf.has_section("book"):  # 检查是否存在section
            cnf.add_section("book")
        if not cnf.has_option("book", "bookfolder"):  # 检查是否存在该option
            cnf.set("book", "bookfolder", dir_book)
            cnf.set("book", "filename", pa_name)
            cnf.set("book", "pages", end)
        cnf.write(open(config_file, "w"))
    
    dir_book = book_path + dir_book   # notice "\\"
    pa_name = pa_name + ".pdf"
    end = int(end)
    print("The pages will be stored in " + dir_book + " and ensure your 1.pdf is also there;")
    print("The merged book will be stored as '" + pa_name + "' in folder " + book_path)
    print("The Book totally contains " + str(end) + " pages.")
    skip = input("Press Enter to continue...\n")

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
            print("Scanning " + png + " ... " + str(tcount), end= "\r")
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
    while tmptime < 50:         # timeout 50*0.8 seconds
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
        time.sleep(0.8)

    time.sleep(0.2)
    pyautogui.typewrite(str(page), interval=0.1)
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'f4')


def getFileName(filedir):
    file_list = [os.path.join(root, filespath) \
                 for root, dirs, files in os.walk(filedir) \
                 for filespath in files \
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []

def MergePDF(filepath, outfile):

    output = PdfFileWriter()
    outputPages = 0
    pdf_fileName = []
    for i in range(1,end+1):
        pdf_fileName.append(filepath+"\\"+str(i)+".pdf")
    if pdf_fileName:
        for pdf_file in pdf_fileName:
            print("File path：%s"%pdf_file)

            input = PdfFileReader(open(pdf_file, "rb"))

            pageCount = input.getNumPages()
            outputPages += pageCount

            # add individual page to output
            for iPage in range(pageCount):
                output.addPage(input.getPage(iPage))

        print("Total pages: %d."%outputPages)
        #outputStream = open(os.path.join(".\\", outfile), "wb")
        outputStream = open(os.path.join(book_path, outfile), "wb")
        output.write(outputStream)
        outputStream.close()
        print("Merge Completed！")

    else:
        print("No file to merge！")

####################################################
############# Main Program #########################

relist = []

print("Welcome to use SAP Learning Hub e-books polling tool. Current version " + version)
print("""
The program is simulating the action of 'print chrome web page to pdf',
so first open the learning hub book, and save the 1st page as '1.pdf' to a local folder,
then execute this program.
""")


Get_input()

print(end)

if is_repoll == "Yes":
    for i in range(start,end+1,1):
        tmp_file = Path(dir_book + "\\"+ str(i)+'.pdf')
        if tmp_file.is_file():
            continue
        else:
            #print(str(i) + " not exist.")
            relist.append(i)
else:                        # new poll
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
            #print("Size of " + filename + ":\t" + str(Size))
            os.remove(fullfile)

print("Individual size of these pages less than 1KB:")
print(relist)
print("Totally " + str(len(relist)) + " slides to re-poll...")

while len(relist):
    tmppage = relist.pop()
    print(tmppage)
    ClickPng(page_input,2)
    pyautogui.typewrite(str(tmppage), interval=0.1)
    pyautogui.press('enter')
    Poll_page(tmppage)

print("Polling Completed!")

for parent,dirnames,filenames in os.walk(dir_book):
    for filename in filenames:
        tmppage = int(filename[:-4])
        fullfile = dir_book + "\\" + filename
        Size=os.path.getsize(fullfile)
        if Size < 1024 and tmppage > 2:
            print("Blank pages still exist, please re-execute the program...")
            sys.exit()

MergePDF(dir_book,pa_name)
os.remove(config_file)
