# Import the main excel file containing the data to be entered to firewall from Desktop

import os
import configparser

import pandas as pd
import string
import random

desktop_path = os.path.join(os.environ["HOMEDRIVE"], os.environ["HOMEPATH"], "Desktop\\")
print('\n')
print(desktop_path)

filename = input("Enter the name of xlsx file in Desktop folder containing the Name and Roll Number (if applicable): \n")
filename = filename.split(".")[0] + '.xlsx'

file_path = os.path.join(desktop_path, filename)

if os.path.exists(file_path):
    print(f"File found!")
else:
    print(f"The file '{filename}' does not exist in '{desktop_path}'")
    print('exiting...')
    exit()

pd.options.mode.chained_assignment = None
data = pd.read_excel(file_path)
print('\n')
print(data)

smallLetters = string.ascii_lowercase
capitalLetters = string.ascii_uppercase
numbers = string.digits
alphabets = smallLetters + capitalLetters + numbers

group = input("\nEnter user group (Students/Faculty): ").lower()
if group not in ['students', 'student','faculty']:
    print("Invalid group name")
    input('Press any key to continue...')
    exit()
val = 1 if 'faculty' in group else 0

possible_names = ["name", "Name", "NAME"]
df_name = [col for col in data.columns if any(names in col for names in possible_names)]
name = data[df_name]
if(isinstance(name.iloc[0].item(), int)):
    print("Invalid names")
    input('Press any key to continue...')
    exit()


try:
    if name[name.columns[0]].isnull().all():
        print("Name column is empty")
        input('Press any key to continue...')
        exit()
except:
    print("Name column does not exists")
    input("Press any key to continue...")
    exit()

try:
    possible_names = ["email", "email address", "Email Address", "Email address", "EMAIL", "EMAIL ADDRESS"]
    df_email = [col for col in data.columns if any(emails in col for emails in possible_names)]
    email = data[df_email]
except:
    pass


try:
    possible_names = ["Roll Number", "roll number", "ROLL NUMBER", "Roll", "roll", "ROLL", "Roll no.", "Roll number:"]
    df_roll = [col for col in data.columns if any(rolls in col for rolls in possible_names)]
    roll = data[df_roll]
except:
    pass


passwords = []
usernames = []
remarks = []


print("\nCreating usernames and generating random passwords...\n")

for i in range(len(name)):
    temp = random.sample(alphabets, random.randint(6, 7))
    password = "".join(temp)
    passwords.append(password)
    if(val):
        uname = name.iloc[i].item().split()[0].lower() + '.' + name.iloc[i].item().split()[-1].lower()
        usernames.append(uname)
    else:
        uname = name.iloc[i].item().split()[0].lower() + '_' + str(roll.iloc[i].item())
        usernames.append(uname)

data['Username'] = usernames
data['Password'] = passwords

# Remove duplicates?
data['Password'][data['Password'].duplicated()] = "".join(random.sample(alphabets, random.randint(6, 7)))

# Create final.xlsx file on Desktop containing the usernames and generated passwords
# resultExcelFile = pd.ExcelWriter(f'{desktop_path}final.xlsx')
# data.to_excel(resultExcelFile, index=False, header=True)
# resultExcelFile.save()

print('\n\nUsernames and Passwords: \n')
print(data)
print('\n\n')
resultExcelFile = pd.ExcelWriter(f'{desktop_path}final.xlsx')
data.to_excel(resultExcelFile, index=False, header=True)
resultExcelFile.save()
# print("final.xlsx file created on Desktop Successfully!")
# exit()

# ******************************* SELENIUM PART ******************************* 


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from time import sleep


config = configparser.ConfigParser()
config.read('credentials.ini')

fortiLink = str(config.get('firewall', 'link'))
if fortiLink.endswith('/'):
    fortiLink = fortiLink[:-1]
fortiUsername = config.get('firewall', 'username')
fortiPassword = config.get('firewall', 'password')


def findXpath(xpath):
    count = 0
    while(1):
        if count > 20:
            print('error...exiting')
            exit()
        try:
            el = driver.find_element(by=By.XPATH, value=xpath)
            if count > 0:
                print('found')
            return el
        except:
            print('error...waiting')
            count += 1
            sleep(2)


userGroup_entries = ['Deerwalk Compware', 
'DSS +2', 
'DWIT Faculty', 
'DWIT Prof. Training', 
'DWIT Students', 
'DWIT_BCA students', 
'Incubation', 
'LIMITED-GROUP',
]

[print(f'{i+1}. {userGroup_entries[i]}') for i in range(len(userGroup_entries))]
gp_index = int(input("\nEnter user group (1/2/3/4/5): ")) - 1

print("\nOpening Selenium...\n")

chrome_options = Options()
chrome_options.add_argument('--ignore-ssl-errors')
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
chrome_options.allow_insecure_localhost = True

try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
except:
    driver = webdriver.Chrome(r"C:\Users\LEGION\Desktop\CODE\DevOps-AddUsers\browsers\chromedriver.exe", options=chrome_options)
driver.maximize_window()

driver.get(fortiLink)

successCount = 0

# Go through the site not secure prompt
try:
    details = driver.find_element(by=By.ID, value="details-button")
    details.click()
    insecure = driver.find_element(by=By.ID, value="proceed-link")
    insecure.click()
except:
    pass

# LOGIN
sleep(3)
username = driver.find_element(by=By.ID, value="username")
# username.send_keys('admin')
username.send_keys(fortiUsername)
password = driver.find_element(by=By.ID, value="secretkey")
# password.send_keys('uvSZgvA5Z')
password.send_keys(fortiPassword)
loginBtn = driver.find_element(by=By.ID, value="login_button")
loginBtn.click()
sleep(7)
print("logged in!\n")
driver.get(f'{fortiLink}/ng/user/local')
sleep(7)


for i in range(len(usernames)):
    driver.get(f'{fortiLink}/ng/user/local')
    sleep(3)
    createNew = findXpath('//*[@id="navbar-view-section"]/div/div/f-local-user-list/f-mutable/div/div[1]/div[2]/div/f-mutable-menu-transclude/f-local-user-list-menu/div[1]/div[1]/div/button')
    createNew.click()
    # sleep(5)

    # nextPg = findXpath('//*[@id="ng-base"]/form/div[3]/dialog-footer/button[2]')
    try:
        nextPg = driver.find_element(by=By.CSS_SELECTOR, value='button[type="submit"]')
    except:
        sleep(7)
        nextPg = driver.find_element(by=By.CSS_SELECTOR, value='button[type="submit"]')
    sleep(1)
    nextPg.click()

    fillName = usernames[i]
    fillPass = passwords[i]

    inpName = findXpath('//*[@id="ng-base"]/form/div[2]/div[1]/div[2]/dialog-content/f-local-user-wizard/section/div[1]/div/input')
    inpName.send_keys(fillName)

    inpPass = findXpath('//*[@id="ng-base"]/form/div[2]/div[1]/div[2]/dialog-content/f-local-user-wizard/section/div[2]/div/input')
    inpPass.send_keys(fillPass)

    
    nextPg.click()
    sleep(1)
    nextPg.click()
    sleep(2)
    # print('heree 0')
    try:
        userGp = driver.find_element(by=By.XPATH, value='//*[@id="ng-base"]/form/div[2]/div[1]/div[2]/dialog-content/f-local-user-wizard/section/div[2]/label[1]/span/label')
        userGp.click()
        remarks.append('User created successfully')
    except:
        remarks.append('User already exists')
        continue
    sleep(2)

    # print('heree 1')
    addUg = findXpath('//*[@id="ng-base"]/form/div[2]/div[1]/div[2]/dialog-content/f-local-user-wizard/section/div[2]/div/div')
    addUg.click()

    # print('heree')
    sleep(2)
    searchEntry = findXpath('//*[@id="navbar-view-section"]/div/div[2]/div/div[2]/div/input')
    searchEntry.send_keys(userGroup_entries[gp_index])
    sleep(1)
    searchEntry.send_keys(Keys.ENTER)
    sleep(2)
    searchEntry.send_keys(Keys.ESCAPE)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    addUg.click()

    # emptyBody = findXpath('//*[@id="ng-base"]/form/div[2]')
    # action = webdriver.ActionChains(driver).move_to_element_with_offset(emptyBody, 20, 50).click().perform()
    # action = webdriver.ActionChains(driver).move_to_element_with_offset(emptyBody, 20, 100).click().perform()
    sleep(2)

    # SUBMIT
    # nextPg.click()
    finalSubmit = findXpath('//*[@id="ng-base"]/form/div[3]/dialog-footer/button[2]')
    finalSubmit.click()

    successCount += 1
    print(f'\n{fillName} activated.')


data['Remarks'] = remarks

# Create final.xlsx file on Desktop containing the usernames and generated passwords
resultExcelFile = pd.ExcelWriter(f'{desktop_path}final.xlsx')
data.to_excel(resultExcelFile, index=False, header=True)
resultExcelFile.save()

print("\n\nfinal.xlsx file created on Desktop Successfully!")
input("Operation completed. Press any key to continue...")

driver.close()