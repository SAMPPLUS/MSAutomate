from selenium import webdriver
import os, glob, shutil
import webbrowser
import time
import xlwings as xw
from datetime import date
import getpass
import json

class Job:
    def __init__(self):
        self.fileName = str()
        self.userName = str()
        self.uId = int()
        self.jobId = int()
        self.downloadLink = str()

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

def CheckExists(user,file):
    pass

def AddToQueue(user, file, path):
    """Adds row to spreadsheet with patron and file names"""
    i = 4
    while True:
        val = str(i)
        cell = sht.range('A' + val)
        if cell.value is None:
            print("queue:", file)
            cell.value = [d1, user, file, None, None, f"=HYPERLINK(\"{path}\", \"{user}'s folder\")"]
            break
        i+= 1



with open("data.json", "r") as dataFile:
    data = json.load(dataFile)

filePath = data["filePath"]
downloadPath = data["downloadPath"]
driverPath = data["driverPath"]

#initialize excel

today = date.today()
d1 = today.strftime("%m/%d/%Y")

wb = xw.Book('example.xlsx')
sht = wb.sheets[0]

#initialize web driver

browser = webdriver.Chrome(executable_path=driverPath)

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

jobList = []

fileNameList = []
userNameList = []
userNameSet = set()

if len(rawInfo) == 0:
    print("no new files.")

for el in rawInfo:
    job = Job()
    print(type(el))
    print(el)
    #file name  
    text = el.text
    fileNameEnd = text.lower().find('.stl') + 4
    fileName = text[:fileNameEnd]
    job.fileName = fileName
    #user name
    userNameStart = text.find('Submitted by') + 13
    userNameEnd = text.find('Price') - 1
    # beware of last name 'Price'
    name = (text[userNameStart:userNameEnd])
    #uid
    uidStart = text.find('&uid=') + 5
    uidEnd = text.find('&job_id=')


    job.userName = name
    userNameSet.add(name)
    jobList.append(job)


#download
downloadLinks = browser.find_elements_by_css_selector("a[aria-label='Download file']")

assert (len(downloadLinks) == len(jobList))

for i, link in enumerate(downloadLinks):
    jobList[i].downloadLink = link


os.chdir(downloadPath)
startSize = len(glob.glob("*.stl"))


print("downloading...")
clickCount = 0
j = 0
for link in downloadLinks:
    link.click()
    clickCount += 1
    time.sleep(1)

    #if CheckExists(userNameList[j],fileNameList[j]):
    #    print("skip: ", fileNameList[j])

    #j += 1
    pass

timer = 0

try:
    # loops until all files have completed downloading
    while len(glob.glob("*.stl")) < (startSize + clickCount):
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

for job in jobList:
    path = filePath + LastFirst(job.userName)


    try:
        AddToQueue(job.userName,job.fileName[:-4], path)

    except Exception as e: 
        print("failed to add to queue " + job.fileName + " - ", e)

    try:
        if not os.path.exists(path):
            print("Making Directory:", path)
            os.mkdir(path)
        shutil.move(job.fileName, path)
        print("moving to directory:", path)

    except Exception as e: 
        print("failed to move " + job.fileName  + " - ", e)


print("file movement complete")

webbrowser.open(filePath)
