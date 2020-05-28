import schedule
import random
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl as excel
try:
    import autoit
except:
    pass
import time
import datetime
import os

browser = None
Contact = None
message = None
Link = "https://web.whatsapp.com/"
wait = None
choice = None
docChoice = None
doc_filename = None
im_filename = None
Ads = None

def whatsapp_login():
    global wait,browser,Link
    browser = webdriver.Chrome()
    wait = WebDriverWait(browser, 600)
    browser.get(Link)
    # browser.maximize_window()
    print("QR scanned")

def sender():
    group_title = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.RLfQR")))
    search = browser.find_elements_by_xpath('//*[@id="side"]/div[1]/div/label/input')[0]

    time.sleep(random.uniform(5, 10))

    for i in Contact:
        print(i)
        send_message(search, i)

def read_message(fileName):
    msgIklan = open(fileName, 'r').read()

    myText = msgIklan[0]
    for j in range(1, len(msgIklan)):
        myText += msgIklan[j]
    myText = myText.replace('\n', (Keys.SHIFT + Keys.ENTER + Keys.SHIFT));

    return myText

def read_Contacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        if contact != 'None':
        # contact = "\"" + contact + "\""
            lst.append(contact)
    return lst

def input_file():
    global Contact, Ads 
    Contact = []
    Ads = []

    while True:
        # x = str(input("Enter database file name: \n"))
        x = 'a.xlsx'
        if os.path.isfile(x) == True:
            Contact = read_Contacts(x)
            # print(Contact)
            break
        elif os.path.isfile(x) == False:
            print('\nNo such file or directory')

    while True:
        # y = str(input("Enter ads file name: \n"))
        y = 'iklan_madu.txt'
        if os.path.isfile(y) == True:
            Ads = read_message(y)
            # print(Ads)
            break
        elif os.path.isfile(y) == False:
            print('\nNo such file or directory')

def send_message(search, target):

    search.clear()
    search.send_keys(target)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "button._3Burg")))
    time.sleep(random.uniform(1, 2))
    persons = browser.find_elements_by_class_name('_2wP_Y')

    for person in persons:
        try:
            if person.text not in ['CHATS','MESSAGES']:
                print('1 ' + person.text)
                print('2 ' + person)
                print('3 ' + persons)
        except:
            print("*")

if __name__ == "__main__":

    input_file()

    # If you want to schedule messages for
    # a particular timing choose yes
    # If no choosed instant message would be sent

    while True:
        # isSchedule = input('\nDo you want to schedule your Message(y/n):')
        isSchedule = 'y'
        if (isSchedule == "y") or (isSchedule == "Y"):
            # jobtime = input('input time in 24 hour (HH:MM) format - ')
            break
        elif (isSchedule == "n") or (isSchedule == "N"):
            # print(' ')
            break

    # #Send Attachment Media only Images/Video
    # choice = input("Would you like to send attachment(yes/no): ")
    # if(choice == "yes"):
    #     # Note the document file should be present in the Document Folder
    #     im_filename = input("Enter the image file name you want to send: ")

    # docChoice = input("Would you file to send a Document file(yes/no): ")
    # if(docChoice == "yes"):
    #     # Note the document file should be present in the Document Folder
    #     doc_filename = input("Enter the Document file name you want to send: ")

    # Let us login and Scan
    print("SCAN YOUR QR CODE FOR WHATSAPP WEB")
    whatsapp_login()

    # # Send message to all Contact List
    # # This sender is just for testing purpose to check script working or not.
    # # Scheduling works below.
    # # sender()
    # # Uncomment line 236 is case you want to test the program

    
    if (isSchedule == "y") or (isSchedule == "Y"):
        # schedule.every().day.at(jobtime).do(sender)
        print(' ')
        sender()
    elif (isSchedule == "n") or isSchedule == "N":
        print(' ')
        sender()


    # # First time message sending Task Complete
    # print("Task Completed")

    # # Messages are scheduled to send
    # # Default schedule to send attachment and greet the personal
    # # For GoodMorning, GoodNight and howareyou wishes
    # # Comment in case you don't want to send wishes or schedule
    # # scheduler()
    
    # # browser.quit()