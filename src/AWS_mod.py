import os
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException

def webloading(browser, link, address_XPATH):    
    browser.get(link)
    timeout = 3 # seconds
    i = 1
    while True:
        try:
            element_present = EC.presence_of_element_located((By.XPATH, address_XPATH))
            WebDriverWait(browser, timeout).until(element_present)
            validation = True
            break
        except TimeoutException:
            try:
                aa = browser.find_element_by_xpath('//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')
                validation = False
                break
            except NoSuchElementException:
                pass
            print (i, "Menunggu halaman web terbuka")
            i+=1
    return validation

def autoitloading(autoit):    
    timeout = 3 # seconds
    i = 1
    while True:
        try:
            autoit.win_wait_active("Open", timeout)
            break
        except TimeoutException:
            print (i, "Menunggu ")
            i+=1

if __name__ == '__main__':
    # wb_name = 'database.xlsx'
    # medialist = read_database(wb_name, 'madu', 'D')


    # image_path = read_medialist(medialist)
    # print(image_path)
    # link_num = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(6282210138809)
    # validatorsnumber(link_num)
    # validatorsnumber(link_num)
    pass