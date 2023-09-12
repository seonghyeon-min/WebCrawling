from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException, WebDriverException
import time
import pyperclip
import pandas as pd

# < --------------------------- precondiion ---------------------------------- > #

CHROMDRIVE_PATH = r"C:\Users\test\Desktop\WebCrawling\chromedriver-win32\chromedriver-win32\chromedriver.exe"  # 반드시 chrome version 확인할 것
USER_ID = 'seonghyeon.min'
USER_PW = 'alstjdgus@3498'
DF = {'PlatformCode' : [], 'Country' : [], 'NPV' : [], 'Status' : [], 'Shelf' : [] } 

# < --------------------------------------------------------------------------------------------- > 


def ProcessCrawling(id, pw) :
    driver = webdriver.Chrome(CHROMDRIVE_PATH) # webdriver는 web을 운전하기 위한 도구라고 보면 됨
    driver.implicitly_wait(3)

    driver.get('http://qt2-kic.smartdesk.lge.com/admin/main.lge?serverType=QA2') # 운전하려는 처음 url 설정 
    
    alert = Alert(driver) 
    alert.accept()
    time.sleep(3)
    
    driver.find_element(By.ID,'USER').click() 
    pyperclip.copy(id)
    driver.find_element(By.ID,'USER').send_keys(Keys.CONTROL,'v')
    
    driver.find_element(By.ID,'LDAPPASSWORD').click()
    pyperclip.copy(pw)
    driver.find_element(By.ID,'LDAPPASSWORD').send_keys(Keys.CONTROL,'v')

    driver.find_element(By.ID,'loginSsobtn').click()
    
    time.sleep(0.5)
    
    driver.get('http://qt2-kic.smartdesk.lge.com/admin/home/shelf/shelfdisplay/retrieveShelfDisplayList.lge?lnbMenu=Y') # home-shelf url 
    driver.find_element(By.XPATH,'//*[@id="sdpForm"]/fieldset/div/table/tbody/tr[1]/td[2]/div').click()

    pyperclip.copy('webOSTV 23-S23Y')
    driver.find_element(By.XPATH, '//*[@id="sdpForm"]/fieldset/div/table/tbody/tr[1]/td[2]/div/div/div/input').send_keys(Keys.CONTROL, 'v')
    driver.find_element(By.XPATH, '//*[@id="sdpForm"]/fieldset/div/table/tbody/tr[1]/td[2]/div/div/div/input').send_keys(Keys.ENTER)

    time.sleep(0.5)
    
    driver.find_element(By.XPATH, '//*[@id="sdpForm"]/fieldset/div/div/button').click()
    
    time.sleep(0.5)
    
    driver.find_element(By.XPATH, '//*[@id="sdpForm"]/div[2]/div[1]/select').click()
    driver.find_element(By.XPATH, '//*[@id="sdpForm"]/div[2]/div[1]/select/option[7]').click()
    
    # <! -- pagination-group -- !> #
    Pagnation = list(driver.find_element(By.XPATH, '//*[@id="sdpForm"]/nav/ul').text)
    PageList = ''.join(Pagnation).split('\n')
    
    try : 
        startpageidx = PageList.index('1')
        endpageidx = PageList.index('Next')-1
    except ValueError :
        startpageidx, endpageidx = 1, 1
        print('> ! -- Pagnation : 1 -- < ')

    for page in range(startpageidx, endpageidx + 1) :
        pageSelector =  f'//*[@id="sdpForm"]/nav/ul/li[{page}]/a'
        driver.find_element(By.XPATH, pageSelector).click()
        time.sleep(0.2)
        
        # < -- check the row of count -- > #
        table = driver.find_element(By.XPATH, '//*[@id="sdpForm"]/div[3]/table/tbody')
        tr = len(table.find_elements(By.TAG_NAME, 'tr'))
        
        print(f' > -- start crawling {tr} of data -- < ')
        
        # < -- click the shelf display ID -- > # 
        for num in range(1, tr+1) :
            time.sleep(0.5)
            shelfDispId = f'//*[@id="sdpForm"]/div[3]/table/tbody/tr[{num}]/td[2]/a'
            driver.find_element(By.XPATH, shelfDispId).send_keys(Keys.ENTER)
            time.sleep(0.5)
            
            ProdPlfCode = driver.find_element(By.XPATH, '//*[@id="sdpForm"]/fieldset/div/table/tbody/tr[2]/td[1]').text
            version = driver.find_element(By.XPATH, '//*[@id="sdpForm"]/fieldset/div/table/tbody/tr[3]/td[1]').text
            country = driver.find_element(By.XPATH, '//*[@id="sdpForm"]/fieldset/div/table/tbody/tr[3]/td[2]').text
            Status = driver.find_element(By.XPATH, '//*[@id="sdpForm"]/fieldset/div/table/tbody/tr[4]/td[1]/span').text
            
            DF['PlatformCode'].append(ProdPlfCode)
            DF['NPV'].append(version)
            DF['Country'].append(country)
            DF['Status'].append(Status)

            table = driver.find_element(By.XPATH, '//*[@id="shlefListTable"]/tbody')
            tr = len(table.find_elements(By.TAG_NAME, 'tr'))
            tmp = []
            
            for row in range(1, tr+1) :
                trpath = f'//*[@id="shlefListTable"]/tbody/tr[{row}]/td[2]'
                shelf = driver.find_element(By.XPATH, trpath).text
                tmp.append(shelf)
            
            DF['Shelf'].append(', '.join(tmp))
                
            driver.back()
    
    print('> -- crawling has been finsihed -- < ')
    df = pd.DataFrame(DF)
    df.to_excel('result.xlsx')
    return DF

ProcessCrawling(USER_ID, USER_PW)
