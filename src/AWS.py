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
from AWS_read import read_message, read_database, read_medialist, read_setting
from AWS_mod import webloading, autoitloading
from AWS_about import about

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

def input_contacts(db_col):
    global unsaved_Contacts
    unsaved_Contacts = []
    unsaved_Contacts = read_database(wb_name, 'database', db_col)
    if len(unsaved_Contacts) != 0:
        print("Daftar nomor yang akan dikirim ->")
        for i in unsaved_Contacts:
            print(i)
    else:
        print('Tidak ada nomor tersedia')

def input_message(campaign):
    global message
    message = read_database(wb_name, campaign, 'A')[0]
    print()
    print('-------------------------------------------')
    print(message)
    print('-------------------------------------------')
    print()
    message = message.split('\n')

def whatsapp_login(chrome_path):
    global browser
    chrome_options = Options()
    chrome_options.add_argument('--user-data-dir=./User_Data')
    # chrome_options.add_argument('--headless')
    browser = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
    wait = WebDriverWait(browser, 600)
    browser.get(Link)
    print("Kode QR telah berhasil di pindai")

def send_message():
    global message
    try: 
        address_XPATH = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
        input_box = browser.find_element_by_xpath(address_XPATH)

        for ch in message:
            if ch == "":
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
    if docType == 1:
        medialist = read_database(wb_name, campaign, 'D')
        medialist_Desc = read_database(wb_name, campaign, 'E')
    else:
        medialist = read_database(wb_name, campaign, 'G')
    
    image_path = read_medialist(medialist, docType)
    print(medialist[1:])    

    clipButton = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div')
    clipButton.click()

    time.sleep(5)
    if docType == 1:
        address_XPATH = '//*[@id="main"]/footer/div[1]/div[1]/div[2]/span/div/div/ul/li[1]/button'
        mediaButton = browser.find_element_by_xpath(address_XPATH)
        mediaButton.click()
    else:
        address_XPATH = '//*[@id="main"]/footer/div[1]/div[1]/div[2]/span/div/div/ul/li[3]/button'
        docButton = browser.find_element_by_xpath(address_XPATH)
        docButton.click()
    autoitloading(autoit)
    time.sleep(1)
    autoit.control_focus("Open", "Edit1")
    autoit.control_set_text("Open", "Edit1", image_path)
    autoit.control_click("Open", "Button1")
    
    if docType == 1:
        time.sleep(5)
        address_XPATH = '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[2]/div/div[3]/div[1]'
        ket = browser.find_element_by_xpath(address_XPATH)   
        ket.send_keys(medialist_Desc[0])

        a = image_path.split(' ')

        for i in range(len(a)-2):
            imPress = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[2]/span/div[' + str(i+2)+ ']')
            imPress.click()

            ket = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[2]/div/div[3]/div[1]')   
            ket.send_keys(medialist_Desc[i+1])


    time.sleep(5)
    whatsapp_send_button = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div')
    whatsapp_send_button.click()

def sendersub(istiboll, istijeda, istiwaktu, i):
    global link_num
    link_num = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(i)
    #driver  = webdriver.Chrome()
    browser.get(link_num)

    address_XPATH = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
    valadation = webloading(browser, link_num, address_XPATH)
    # input_box = browser.find_element_by_xpath(address_XPATH)

    if valadation == True:
        print("Mengirim pesan ke", i)
        send_message()
        print('Pesan terkirim')
        if(choice == "yes"):
            try:
                docType = 1
                send_attachment(docType)
                print('Gambar/video terkirim')
            except:
                print('Gambar/video tidak terkirim')
        if(docChoice == "yes"):
            try:
                docType = 2
                send_attachment(docType)
                print('Dokumen terkirim')
            except:
                print('Dokumen tidak terkirim')
    time.sleep(1)

def sender(istiboll, istijeda, istiwaktu):
    # global link_num
    if len(unsaved_Contacts) > 0:
        for i in unsaved_Contacts: # for i in range(1, 100):
            if istiboll == 'yes':
                if (int(i) % istijeda != 0):
                	########
                	sendersub(istiboll, istijeda, istiwaktu, i)
                    ########
                else:
                    print('Istirahat dulu bos,', istiwaktu, 'menit')
                    time.sleep(float(istiwaktu)*60)
            else:
                sendersub(istiboll, istijeda, istiwaktu, i)
    else:
        print('Tidak ada nomor tersedia')


if __name__ == "__main__"

    about()

    wb_name = 'database.xlsx'
    db_col, campaign, choice, docChoice, istiboll, istijeda, istiwaktu = read_setting(wb_name)

    input_contacts(db_col)
    input_message(campaign)

    print("Pindai kode QR")
    whatsapp_login(args.chrome_driver_path)

    sender(istiboll, istijeda, istiwaktu)

    print("Tugas selesai")
    
    # browser.quit()
