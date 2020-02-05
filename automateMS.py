from selenium import webdriver
from pathlib import Path
import os, glob, shutil
import webbrowser
import time
import xlwings as xw
from datetime import date
import getpass
import json
import re

class Job:
    def __init__(self):
        self.fileName = str()
        self.userName = str()
        self.uId = int()
        self.jobId = int()
        self.downloadLink = str()
        self.infoLink = str()
        self.skip = True
        self.show = True

    def __repr__(self):
        rep = f"{self.fileName} - {self.userName}"
        return rep

clear = lambda: os.system('cls')

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

def AddToQueue(user, file, path, infoLink):
    """Adds row to spreadsheet with patron and file names"""
    i = 4
    while True:
        val = str(i)
        cell = sht.range('A' + val)
        if cell.value is None:
            print("queue:", file)
            cell.value = [d1, user, file, f"=HYPERLINK(\"{infoLink}\", \"Job Info\")", f"=HYPERLINK(\"{path}\", \"{user}'s folder\")" ]
            break
        i+= 1

def DisplayJobs(jobList, useSet, status):
    clear()
    print(status, '\n')

    print(len(jobList), "AVAILABLE JOBS:")
    for i, job in enumerate(jobList):
        print(f"[{i+1}] - {job}")
    
    print("\n[e] - execute with selected jobs")
    print("[a] - execute all")
    if len(useSet) == 0:
        useSet = "{}"
    print("selected jobs:", useSet)

def ChooseJobs(jobList):
    visibleList = []
    for job in jobList:
        if job.show:
            visibleList.append(job)

    status = "Choose how to proceed"
    useSet = set()

    while True:
        DisplayJobs(visibleList, useSet, status)
        user_input = input()

        if user_input.isalpha():

            if user_input.lower() == 'e':
                if len(useSet) < 1:
                    status = "please select at least one job"
                else:
                    status = "executing " + useSet.__repr__()
                    for el in useSet:
                        visibleList[el-1].skip = False
                    break


            elif user_input.lower() == 'a':
                status = "executing all"
                for job in visibleList:
                    job.skip = False
                break

            else:
                status = "unrecognized command"

        elif user_input.isdecimal():

            if (int(user_input) < 1) or (int(user_input) > len(visibleList)):
                status = ("Error - Out of Range")

            else:

                if int(user_input) in useSet:
                    useSet.remove(int(user_input))
                    status = "removed " + (user_input)

                else:
                    useSet.add(int(user_input))
                    status ="added " + (user_input)
        else:
            status = "unrecognized command"
    
    clear()
    print(status)
    return visibleList


# read data

with open("data.json", "r") as dataFile:
    data = json.load(dataFile)

filePath = Path(data["filePath"])
downloadPath = Path(data["downloadPath"])
driverPath = Path(data["driverPath"])
queuePath = data["queuePath"]

#initialize web driver
print(filePath.__repr__())
options = webdriver.ChromeOptions() 
prefs = {'download.default_directory' : str(downloadPath)}
options.add_experimental_option('prefs', prefs)

browser = webdriver.Chrome(options=options)

browser.get('https://3dprime.lib.msu.edu/')

#Login
loggedIn = 0

if not loggedIn:
    emailElem = browser.find_element_by_name('email')
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
    #file name
    text = el.text
    StlPattern = re.compile(".*\.(stl)")
    match = re.search(StlPattern, text.lower())
    if match:
        fileName = match.group(0)
        job.fileName = fileName
    else:
        job.show = False
        job.skip = True
    #user name
    userNameStart = text.find('Submitted by') + 13
    userNameEnd = text.find('Price') - 1
    # beware of last name 'Price'
    name = (text[userNameStart:userNameEnd])
    job.userName = name
    userNameSet.add(name)
    #uid
    uidStart = text.find('&uid=') + 5
    uidEnd = text.find('&job_id=')


    jobList.append(job)



#info
infoLinks = browser.find_elements_by_css_selector("a[class='btn btn-dark prime_float_right']")
for i, link in enumerate(infoLinks):
    jobList[i].infoLink = link.get_attribute("href")

#download
downloadLinks = browser.find_elements_by_css_selector("a[aria-label='Download file']")

assert (len(downloadLinks) == len(jobList))

for i, link in enumerate(downloadLinks):
    jobList[i].downloadLink = link.get_attribute("href")

jobList = ChooseJobs(jobList)

os.chdir(downloadPath)
files = glob.glob("*")
for f in files:
    os.remove(f)

startSize = len(glob.glob("*.stl"))

print("downloading...")
clickCount = 0
j = 0
for job in jobList:
    if job.skip:
        continue
    print("getting", job.downloadLink)
    browser.get(job.downloadLink)
    clickCount += 1

    #if CheckExists(userNameList[j],fileNameList[j]):
    #    print("skip: ", fileNameList[j])

    #j += 1
    pass

timer = 0

#print("start size:", startSize)
#print("click count:", clickCount)

try:
    # loops until all files have completed downloading
    while len(glob.glob("*.stl")) < (startSize + clickCount):
        if timer == 9:
            print("this is taking a while...(press ctrl+c to continue anyway)", f"[{len(glob.glob('*.stl'))}/{(startSize + clickCount)}]")
        time.sleep(1)
        timer += 1
except KeyboardInterrupt:
    pass
    
    
print("downloads complete")
#rename to disinclude 3DP added numbers
for file in glob.glob("*.stl"):
    shutil.move(file,file[5:])


#initialize excel

today = date.today()
d1 = today.strftime("%m/%d/%Y")

wb = xw.Book(queuePath)
sht = wb.sheets[0]

#move files to proper directory and update queue
i = 0  

for job in jobList:
    if job.skip:
        continue
    path = filePath / LastFirst(job.userName)
    try:
        AddToQueue(job.userName,job.fileName[:-4], path, job.infoLink)

    except Exception as e: 
        print("failed to add to queue " + job.fileName + " - ", e)

    try:
        if not path.exists():
            print("Creating Directory:", path)
            os.mkdir(path)
        shutil.move(job.fileName, path)
        print("moving to directory:", path)

    except Exception as e: 
        print("failed to move " + job.fileName  + " - ", e)


print("file movement complete. I am done.")
browser.quit()
webbrowser.open(filePath)
done = input()
