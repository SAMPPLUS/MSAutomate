from selenium import webdriver
import os, glob, shutil
import webbrowser
import time
import xlwings as xw
from datetime import date
import getpass

def LastFirst(name):
    """Flips name into LAST, FIRST or LAST, FIRST MIDDLE format. No change to single-word names.
     ex: 'John Doe' -> 'Doe, John'... 'Jane Marie Doe' -> 'Doe, Jane Marie'... 'Kevin' -> 'Kevin'."""
    nameList = name.split()
    if len(nameList) > 1:
        nameStr = nameList[-1] + ','
        i = 0
        while i < (len(nameList) - 1):
            nameStr += ' ' + nameList[i]
            i += 1
        return nameStr

    return name

def AddToQueue(user, file):
    """Adds row to spreadsheet with patron and file names"""
    i = 4
    while True:
        val = str(i)
        cell = sht.range('A' + val)
        if cell.value is None:
            cell.value = [d1, user, file]
            break
        i+= 1


#initialize excel

today = date.today()
d1 = today.strftime("%m/%d/%Y")

wb = xw.Book('example.xlsx')
sht = wb.sheets[0]

#initialize web driver

browser = webdriver.Chrome()

browser.get('https://3dprime.lib.msu.edu/')

#Login
loggedIn = 0
emailElem = browser.find_element_by_name('email')

if not loggedIn:
    email = input("email: ")
    password = getpass.getpass()
    emailElem.send_keys(email)
    passwordElem =  browser.find_element_by_name('signin_password')
    passwordElem.send_keys(password)
    passwordElem.submit()

browser.get('https://3dprime.lib.msu.edu/?t=all_jobs')

#parse

rawInfo =  browser.find_elements_by_css_selector("div[class='card prime_long_card mb-9']")

fileNameList = []
userNameList = []
userNameSet = set()

if len(rawInfo) == 0:
    print("no new files.")

for el in rawInfo:
    #file name
    text = el.text
    ind = text.lower().find('.stl') + 4
    fileName = text[:ind]
    fileNameList.append(fileName)
    #user name
    st = text.find('Submitted by') + 13
    en = text.find('Price') - 1
    #TODO: beware of last name 'Price'
    name = (text[st:en])

    userNameList.append(name)
    userNameSet.add(name)


#download
downloadLinks = browser.find_elements_by_css_selector("a[aria-label='Download file']")

os.chdir(r"C:\\Users\samdp\Downloads")

startSz = len(glob.glob("*.stl"))

print("downloading...")
clickCount = 0
j = 0
for link in downloadLinks:
    if userNameList[j] == "Sam Peterson":
        link.click()
        clickCount += 1
        #time.sleep(1)
    j += 1

timer = 0
try:
    while len(glob.glob("*.stl")) < (startSz + clickCount):
        if timer == 10:
            print("this is taking a while...(press ctrl+c to continue anyway)")
        time.sleep(1)
        timer += 1
except KeyboardInterrupt:
    pass
    
    
print("downloads complete")

#rename to disinclude 3DP added numbers
for file in glob.glob("*.stl"):
    shutil.move(file,file[5:])

#move files to proper directory and update queue
i = 0
while i < len(fileNameList):
    try:
        AddToQueue(userNameList[i],fileNameList[i][:-4])
    except Exception as e: 
        print(e)
        print("failed to add to queue " + fileNameList[i])
    try:
        if not os.path.exists(r"C:\\Users\samdp\Desktop\\Projects\\MakerSpaceAutomation\August 2019\\" + LastFirst(userNameList[i])):
            os.mkdir(r"C:\\Users\samdp\Desktop\\Projects\\MakerSpaceAutomation\August 2019\\" + LastFirst(userNameList[i]))
        shutil.move(fileNameList[i], r"C:\\Users\samdp\Desktop\\Projects\\MakerSpaceAutomation\August 2019\\" + LastFirst(userNameList[i]))
    except Exception as e: 
        print(e)
        print("failed to move " + fileNameList[i])

    i+= 1

print("file movement complete")

webbrowser.open(r"C:\\Users\samdp\Desktop\\Projects\\MakerSpaceAutomation\August 2019")
