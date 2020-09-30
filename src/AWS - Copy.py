import schedule

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

try:
    import autoit
except ModuleNotFoundError:
    print('ModuleNotFoundError')
    pass
import time
import datetime
import os
import argparse
from AWSmodule import read_message, read_database, read_medialist, read_setting, webloading, autoitloading

parser = argparse.ArgumentParser(description='PyWhatsapp Guide')
parser.add_argument('--chrome_driver_path', action='store', type=str, default='./chromedriver.exe', help='chromedriver executable path (MAC and Windows path would be different)')
parser.add_argument('--message', action='store', type=str, default='', help='Enter the msg you want to send')
parser.add_argument('--remove_cache', action='store', type=str, default='False', help='Remove Cache | Scan QR again or Not')
args = parser.parse_args()

if args.remove_cache == 'True':
    os.system('rm -rf User_Data/*')
browser = None
# Contact = None
message = None if args.message == '' else args.message
Link = "https://web.whatsapp.com/"
link_num = None
wait = None
choice = None
docChoice = None
unsaved_Contacts = None
wb_name = None

# def clearTerminal():
#     clear = lambda: os.system('cls')
#     clear()

def input_contacts(db_col):
    # global Contact, unsaved_Contacts
    global unsaved_Contacts
    # List of Contacts
    # Contact = []
    unsaved_Contacts = []
    # while True:
        # while True:
        #     try:
        #         print("PLEASE CHOOSE ONE OF THE OPTIONS:")
        #         print("1. Message to Saved Contact number")
        #         print("2. Message to Unsaved Contact number")
        #         print("3. Message to Saved Contact on database")
        #         print("4. Message to Unsaved Contact on database")
        #         print('-------------------------------------------')
        #         # x = int(input("Enter your choice(1, 2, 3, or 4):\n"))
        #         x = 4
        #         break
        #     except:
        #         print("That's not a valid option!")
        #         clearTerminal()
        # x = 4
        # if x == 1 or x == 2 or x == 3 or x == 4:
            # if x == 1:
            #     n = int(input('Enter number of Contacts to add(count)->'))
            #     print()
            #     for i in range(0,n):
            #         inp = str(input("Enter contact name(text)->"))
            #         inp = '"' + inp + '"'
            #         # print (inp)
            #         Contact.append(inp)
            # elif x == 2:
            #     n = int(input('Enter number of unsaved Contacts to add(count)->'))
            #     print()
            #     for i in range(0,n):
            #         # Example use: 919899123456, Don't use: +919899123456
            #         # Reference : https://faq.whatsapp.com/en/android/26000030/
            #         inp = str(input("Enter unsaved contact number with country code(interger):\n\nValid input: 91943xxxxx12\nInvalid input: +91943xxxxx12\n\n"))
            #         # print (inp)
            #         unsaved_Contacts.append(inp)
            # elif x == 3:
            #     while True:
            #         db_contact = str(input("Enter database file name : "))
            #         if os.path.isfile(db_contact) == True:
            #             Contact = readContacts3(db_contact)
            #             break
            #         elif os.path.isfile(db_contact) == False:
            #             print('\nNo such file or directory')
            # elif x == 4:
            # if x == 4:
            #     while True:
            #         # db_contact = str(input("Enter database file name : "))
            #         db = file
            #         if os.path.isfile(db) == True:
    unsaved_Contacts = read_database(wb_name, 'database', db_col)
                        # break
        #             elif os.path.isfile(db) == False:
        #                 print('\nNo such file or directory')
        #     # choi = input("Do you want to add more contacts(y/n)->")
        #     # if choi == "n":
        #     break
        # else:
        #     print("That's not a valid option!")
        #     clearTerminal()

    # if len(Contact) != 0:
    #     print("\nSaved contacts entered list->",Contact)
    if len(unsaved_Contacts) != 0:
        print("Daftar nomor yang akan dikirim ->")
        for i in unsaved_Contacts:
            print(i)
    else:
        print('Tidak ada nomor tersedia')
    # input("\nPress ENTER to continue...")

def input_message(campaign):
    global message
    # Enter your Good Morning Msg

    # while True:
        # while True:
        #     try:
        #         print('-------------------------------------------')
        #         print("PLEASE CHOOSE ONE OF THE OPTIONS:")
        #         print("1. Type the message manually")
        #         print("2. Read the message from a file")
        #         print('-------------------------------------------')
        #         # x = int(input("Enter your choice(1 or 2):\n"))
        #         x = 2
        #         break
        #     except:
        #         print("That's not a valid option!")
        #         clearTerminal()
        # x = 2
        # if x == 1 or x == 2 :
            # if x == 1:
            #     # print()
            #     print("Enter the message and use the symbol '~' to end the message:\nFor example: Hi, this is a test message~\n\nYour message: ")
            #     message = []
            #     temp = ""
            #     done = False

            #     while not done:
            #         temp = input()
            #         if len(temp)!=0 and temp[-1] == "~":
            #             done = True
            #             message.append(temp[:-1])
            #         else:
            #             message.append(temp)
            #     message = "\n".join(message)
            #     print()
            #     print(message)
            # elif x == 2:
            # if x == 2:
            #     # print()                
            #     while True:
            #         # y = str(input("Enter message file name: \n"))
            #         y = file
            #         if os.path.isfile(y) == True:
            #             # message = read_message(y)
    message = read_database(wb_name, campaign, 'A')[0]
                    #     break
                    # elif os.path.isfile(y) == False:
                    #     print('\nNo text file or directory')

    print()
    print('-------------------------------------------')
    print(message)
    print('-------------------------------------------')
    print()
    message = message.split('\n')
        #     break
        # else:
        #     print("That's not a valid option!")
        #     clearTerminal()

def whatsapp_login(chrome_path):
    # global wait, browser, Link
    global browser

    chrome_options = Options()
    chrome_options.add_argument('--user-data-dir=./User_Data')
    browser = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
    wait = WebDriverWait(browser, 600)
    browser.get(Link)
    # browser.maximize_window()
    print("Kode QR telah berhasil di pindah")

# def send_message(target):
#     global message, wait, browser
#     try:
#         x_arg = '//span[contains(@title,' + target + ')]'
#         ct = 0
#         while ct != 10:
#             try:
#                 group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
#                 group_title.click()
#                 break
#             except:
#                 ct += 1
#                 time.sleep(10)
#         input_box = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
#         for ch in message:
#             if ch == "\n":
#                 ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
#             else:
#                 input_box.send_keys(ch)
#         input_box.send_keys(Keys.ENTER)
#         print("Message sent successfuly")
#         time.sleep(10)
#     except NoSuchElementException:
#         return

def send_unsaved_contact_message():
    global message
    try:
        # time.sleep(15)
        
        address_XPATH = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
        # webloading(browser, link_num, address_XPATH)
        input_box = browser.find_element_by_xpath(address_XPATH)
        
        # for ch in message:
        #     if ch == "\n":
        #         
        #     else:
        #         input_box.send_keys(ch)
        # # input_box.send_keys(message)

        # message = message.split('\n')
        # print(message)

        for ch in message:
            # ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
            if ch == "":
                # ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
                # time.sleep(1)
                ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
            else:
                ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
                input_box.send_keys(ch)
        # time.sleep(5)
        input_box.send_keys(Keys.ENTER)

        # print("Pesan berhasil dikirimkan")
    except NoSuchElementException:
        print("Pesan gagal dikirimkan")
        return

def send_attachment(docType):
    # global im_filename
    # medialist = open('madu_media.txt', 'r').read().split('\n')
    # docType = 1
    if docType == 1:
        medialist = read_database(wb_name, campaign, 'D')
        medialist_Desc = read_database(wb_name, campaign, 'E')
    else:
        medialist = read_database(wb_name, campaign, 'G')

    # print('aaa', medialist)
    
    image_path = read_medialist(medialist, docType)
    print(medialist[1:])    

    # for i in range(len(medialist)):
    # Attachment Drop Down Menu
    # clipButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[3]/div/div[2]/div/span')
    clipButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div')
    clipButton.click()

    # To send Videos and Images.
    time.sleep(5)
    if docType == 1:
        address_XPATH = '//*[@id="main"]/footer/div[1]/div[1]/div[2]/span/div/div/ul/li[1]/button'
        mediaButton = browser.find_element_by_xpath(address_XPATH)
        mediaButton.click()
    # mediaButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[3]/div/div[2]/span/div/div/ul/li[1]/button')
    # webloading(browser, link_num, address_XPATH)
    else:
        address_XPATH = '//*[@id="main"]/footer/div[1]/div[1]/div[2]/span/div/div/ul/li[3]/button'
        docButton = browser.find_element_by_xpath(address_XPATH)
        docButton.click()
    # image_path = os.getcwd() + "\\Media\\" + 'madu(1).jpeg'
    # image_path = image_path.replace('\\src', '')
    # print(image_path)

    # time.sleep(5)
    autoitloading(autoit)
    time.sleep(1)
    autoit.control_focus("Open", "Edit1")
    autoit.control_set_text("Open", "Edit1", image_path)
    autoit.control_click("Open", "Button1")
    
    if docType == 1:
        time.sleep(5)
        address_XPATH = '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[2]/div/div[3]/div[1]'
        # webloading(browser, link_num, address_XPATH)
        ket = browser.find_element_by_xpath(address_XPATH)   
        ket.send_keys(medialist_Desc[0])

        # time.sleep(5)
        a = image_path.split(' ')

        for i in range(len(a)-2):
            imPress = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[2]/span/div[' + str(i+2)+ ']')
            imPress.click()
            # print(i+2)
            ket = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[2]/div/div[3]/div[1]')   
            ket.send_keys(medialist_Desc[i+1])
            # print(medialist_Desc[i+1])

    time.sleep(5)
    # whatsapp_send_button = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span[2]/div/div/span')
    whatsapp_send_button = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div')
    whatsapp_send_button.click()

# def send_files():
#     #Function to send Documents(PDF, Word file, PPT, etc.)
#     # global doc_filename
#     # filelist = open(doc_filename, 'r').read().split('\n')
#     # print(filelist)

#     # global im_filename
#     # medialist = open('madu_media.txt', 'r').read().split('\n')
#     medialist = read_database(wb_name, campaign, 'G')
#     # medialist_Desc = read_database(wb_name, campaign, 'H')
#     # print('aaa', medialist)
#     # docType = 2
#     image_path = read_medialist(medialist, docType)
#     print(medialist[1:])

#     # for i in range(len(filelist)):
#     # # Attachment Drop Down Menu
#     #     # clipButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[3]/div/div[2]/div/span')
#     clipButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div')
#     clipButton.click()

#     time.sleep(5)
#     #     # To send a Document(PDF, Word file, PPT)
#     #     # docButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[3]/div/div[2]/span/div/div/ul/li[3]/button')
#     docButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/span/div/div/ul/li[3]/button')
#     docButton.click()

#     # time.sleep(5)
#     #     docPath = os.getcwd() + "\\Documents\\" + 'FAQ Madu.pdf'
#     #     docPath = docPath.replace('\\src', '')
#     autoitloading(autoit)
#     time.sleep(1)
#     autoit.control_focus("Open", "Edit1")
#     autoit.control_set_text("Open", "Edit1", image_path)
#     autoit.control_click("Open", "Button1")

#     # time.sleep(5)
#     # ket = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[2]/div/div[3]/div[1]')   
#     # ket.send_keys(medialist_Desc[0])

#     time.sleep(5)
#     # whatsapp_send_button = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span[2]/div/div/span')
#     whatsapp_send_button = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div')
#     whatsapp_send_button.click()

def sender(istiboll, istijeda, istiwaktu):
    # global Contact, choice, docChoice, unsaved_Contacts
    # global choice, docChoice, unsaved_Contacts
    global link_num
    # for i in Contact:
    #     send_message(i)
    #     print("Message sent to ", i)
    #     if(choice == "yes"):
    #         try:
    #             send_attachment()
    #         except:
    #             print('Attachment not sent.')
    #     if(docChoice == "yes"):
    #         try:
    #             send_files()
    #         except:
    #             print('Files not sent')
    # time.sleep(10)
    if len(unsaved_Contacts) > 0:
        for i in unsaved_Contacts: # for i in range(1, 100):
            if istiboll == 'yes':
                if (int(i) % istijeda != 0):
                    link_num = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(i)
                    #driver  = webdriver.Chrome()
                    browser.get(link_num)

                    address_XPATH = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
                    valadation = webloading(browser, link_num, address_XPATH)
                    # input_box = browser.find_element_by_xpath(address_XPATH)

                    if valadation == True:
                        print("Mengirim pesan ke", i)
                        # send_unsaved_contact_message()
                        print('Pesan terkirim')
                        if(choice == "yes"):
                            try:
                                docType = 1
                                # send_attachment(docType)
                                print('Gambar/video terkirim')
                            except:
                                print('Gambar/video tidak terkirim')
                        if(docChoice == "yes"):
                            try:
                                # send_files()
                                docType = 2
                                # send_attachment(docType)
                                print('Dokumen terkirim')
                            except:
                                print('Dokumen tidak terkirim')
                    time.sleep(1)
                else:
                    print('Istirahat dulu bos,', istiwaktu, 'menit')
                    time.sleep(float(istiwaktu)*60)
            else:
                print('a')
    else:
        print('Tidak ada nomor tersedia')

# For GoodMorning Image and Message
# schedule.every().day.at("07:00").do( sender )
# # For How are you message
# schedule.every().day.at("13:35").do( sender )
# # For GoodNight Image and Message
# schedule.every().day.at("22:00").do( sender )
# # Example Schedule for a particular day of week Monday
# schedule.every().monday.at("08:00").do(sender)


# To schedule your msgs
# def scheduler():
#     while True:
#         schedule.run_pending()
#         time.sleep(10)

if __name__ == "__main__":

    # global wb_name

    # clearTerminal()
    print('''
---------------------------------------------
Automatic Whatsapp Sender (AWS)
Created by arif.darmawan@riflab.com
---------------------------------------------
version 1.0
---------------------------------------------
''')


    wb_name = 'database.xlsx'
    db_col, campaign, choice, docChoice, istiboll, istijeda, istiwaktu = read_setting(wb_name)
    # Fmedia, Fdocument = read_setting('setting.txt')

    # Append more contact as input to send messages
    input_contacts(db_col)
    # Enter the message you want to send
    input_message(campaign)

    # If you want to schedule messages for
    # a particular timing choose yes
    # If no choosed instant message would be sent
    # isSchedule = input('Do you want to schedule your Message (y/n):')
    # if (isSchedule=="y") or (isSchedule=="Y"):
    #     jobtime = input('input time in 24 hour (HH:MM) format - ')

    #Send Attachment Media only Images/Video
    # choice = input("Would you like to send media (image or video) file (y/n): ")
    # if Fmedia != '-':
    # # if (choice == "y") or (choice == "Y"):
    #     choice = 'yes'
    #     while True:
    #     # Note the document file should be present in the Document Folder
    #         # im_filename = input("Enter the Media list name you want to send: ")
    #         im_filename = Fmedia
    #         # db_contact = str(input("Enter database file name : "))
    #         if os.path.isfile(im_filename) == True:
    #             # im_filename = (im_filename)
    #             im_filename = Fmedia
    #             break
    #         elif os.path.isfile(im_filename) == False:
    #             print('\nNo such file or directory')
    # # docChoice = input("Would you file to send a document file(y/n): ")
    # if Fdocument != '-':
    # # if (docChoice == "y") or (docChoice == "Y"):
    #     docChoice = 'yes'
    #     while True:
    #     # Note the document file should be present in the Document Folder
    #         # doc_filename = input("Enter the Document list name you want to send: ")
    #         doc_filename = Fdocument
    #         # db_contact = str(input("Enter database file name : "))
    #         if os.path.isfile(doc_filename) == True:
    #             # im_filename = (im_filename)
    #             doc_filename = Fdocument
    #             break
    #         elif os.path.isfile(doc_filename) == False:
    #             print('\nNo such file or directory')

    # Let us login and Scan
    print("Pindai kode QR")
    whatsapp_login(args.chrome_driver_path)

    # Send message to all Contact List
    # This sender is just for testing purpose to check script working or not.
    # Scheduling works below.
    # sender()
    # Uncomment line 236 is case you want to test the program

    # if(isSchedule=="y") or (isSchedule=="Y"):
    #     schedule.every().day.at(jobtime).do(sender)
    # else:

    sender(istiboll, istijeda, istiwaktu)

    # First time message sending Task Complete
    print("Tugas selesai")

    # Messages are scheduled to send
    # Default schedule to send attachment and greet the personal
    # For GoodMorning, GoodNight and howareyou wishes
    # Comment in case you don't want to send wishes or schedule
    # scheduler()
    
    # browser.quit()
