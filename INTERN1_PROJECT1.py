from datetime import datetime
import time
from time import sleep
import pandas as pd
import socket
import os
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from datetime import date
import pyodbc
from selenium import webdriver
from selenium.webdriver.support.ui import Select
#from selenium.common.exceptions import NoSuchElementException
#import http.client
#from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.support.expected_conditions import element_to_be_clickable


def is_internet_available():
    try:
        # see if we can resolve the host name -- tells us if there is
        # a DNS listening
        #host = socket.gethostbyname(hostname)
        # connect to the host -- tells us if the host is actually
        # reachable
        IPaddress=socket.gethostbyname(socket.gethostname())
        if (IPaddress !="127.0.0.1"):
            return True
    except:
        pass
    return False

def clear_cache():
     driver.get('chrome://settings/clearBrowserData')
     sleep(10)
     driver.find_element_by_xpath("/html/body/settings-ui//div[2]/settings-main//settings-basic-page//div[1]/settings-section[4]/settings-privacy-page//settings-animated-pages/div/cr-link-row[1]//cr-icon-button//div/iron-icon").click()
     wait=WebDriverWait(driver,10)
     d = driver.find_element_by_id("clearBrowsingDataConfirm")
     #wait.until(EC.element_to_be_clickable((d))
     wait.until(EC.element_to_be_clickable((d))).click()
     #driver.find_element_by_id("clearBrowsingDataConfirm").click()

flag=1
oldrev=127
d1="16-03-2021"
df1=pd.read_excel(r"ORIGINAL FILE PATH",skiprows=4)
Dflist=list()

while(flag==1):
    if(is_internet_available()):
        try:
            Dflist.clear()
            chromeOptions=webdriver.ChromeOptions()
            prefs = {"download.default_directory": r"path where your files will be stored after loction"}
            chromeOptions.add_experimental_option("prefs", prefs)
            chromedriver="C:\DRIVERS\WIN\chromedriver_win3288v\chromedriver.exe"
            driver=webdriver.Chrome(executable_path=chromedriver,chrome_options=chromeOptions)

            driver.get("WEBSITE ADDRESS")
            print(driver.title)

            driver.find_element_by_xpath("//*[@id='byDetails']").click()

            drp1=Select(driver.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[2]/div[1]/div/div[2]/div/div[1]/div[2]/select"))
            sleep(10)
            drp1.select_by_visible_text('BALCO')
            driver.find_element_by_id("SubmitBtn").click()
            sleep(10)

            #dnld=driver.find_element_by_class_name("icon-download-alt white"))
            #dnld.select_by_visible_text("CSV")
            divtable = driver.find_element_by_class_name("widget-main")
            #print(divtable.text)
            #print("divtable done")

            driver.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[2]/a/i").click()
            sleep(2)

            revino = Select(driver.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[2]/div[1]/div/div[2]/div/div[2]/div[2]/select"))

            item = revino.first_selected_option.text
            print("Rev is ", item)
            print("old rev is", oldrev)

            today = date.today()
            d2 = today.strftime("%d-%m-%Y")
            #print(d2)


            #print(df2)

            #COMPARING DATASHEETS
            if(d1==d2 and is_internet_available()):
                if(item==oldrev and is_internet_available()):
                     #print(df2)
                     print("Found equal")
                     driver.close()
                     driver.quit()
                     print("Wanna continue? Y/N")
                     k=input()
                     flag=(k=='Y') and 1 or 0
                else:
                     driver.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[2]/ul/li[2]/a").click()
                     sleep(1)
                     path="**********************NetSchedule-Summary-BALCO("+item+")-"+d2+".xlsx"            #path to store the newwest downloaded file
                     df2=pd.read_excel(path,skiprows=4)
                     print(df2.shape)
                     oldrev=item

                     if(df1.equals(df2) and is_internet_available()):
                         df1=df2
                         sleep(2)
                         path="*************"              #path where your all files are downloaded
                         files = os.listdir(path)
                         for fil in files:
                             print(fil)
                             os.remove(path+'\\'+fil)
                             print("\tremoved")
                             time.sleep(5)


                     else:

                         diffdf=pd.concat([df1,df2],ignore_index=True, sort=False).drop_duplicates(keep=False)
                         df1=df2
                         print("Dataframe changed")
                         print(diffdf)
                         print(diffdf.shape)
                         firstwocol=diffdf.iloc[:,1:2]
                         print(firstwocol)
                         uniqueno=firstwocol['Time Desc'].unique()
                         print(uniqueno)
                         #uniquedf=pd.Dataframe(uniqueno)
                         uniquestr=uniqueno.tolist()
                         uniquestr.pop()
                         uniquestr.pop()
                         uniquestr.pop()
                         uniquestr.pop()
                         print(uniquestr)
                         uniqueliststr='], ['.join(str(x) for x in uniquestr)
                         print(uniqueliststr)
                         uniquetime=str(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                         print(uniquetime)

                         conn = pyodbc.connect(r'DRIVER={ODBC Driver 17 for SQL Server};SERVER=**SERVERNAME**;DATABASE=DATABASENAME;Trusted_Connection=yes')
                         cursor = conn.cursor()

                         cursor.execute("INSERT INTO DbBALCO(CHANGED_DETAIL, REVNO, TIMESTAMP) values(?,?,?)","["+uniqueliststr+"]",item,uniquetime)
                         cursor.execute("SELECT * from DbBALCO")

                         for i in cursor:
                             print(i)

                         conn.commit()
                         cursor.close()

                         driver.close()
                         driver.quit()


            else:
                d1=d2
                driver.close()
                driver.quit()

        finally:
            if(flag):
                print("NEXT ROUND")
                pass
            else:
                print("BYE!!\nTHANK YOU!!")
                break

    else:
        print("NET NOT AVAILABLE\nKindly check your net connection!")
        flag=0