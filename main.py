from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException
import xlsxwriter

trafficWorkbook=xlsxwriter.Workbook("Traffic record.xlsx")
trafficSheet=trafficWorkbook.add_worksheet()

URL="https://www.google.com/maps/dir/1239+FL-436+%23101,+Casselberry,+FL+32707/2101+Water+Bridge+Blvd,+Orlando,+FL+32837/@28.5319064,-81.517875,11z/data=!3m1!4b1!4m13!4m12!1m5!1m1!1s0x88e76e36b22d92eb:0x4c749fb2c5492ea5!2m2!1d-81.3231343!2d28.6347265!1m5!1m1!1s0x88e77d8ac5cac91b:0x10c348dd5c0c709a!2m2!1d-81.4050814!2d28.4034495"
time_spent_string="#section-directions-trip-0 > div > div:nth-child(1) > div.xB1mrd-T3iPGc-iSfDt-n5AaSd > div.xB1mrd-T3iPGc-iSfDt-duration.delay-heavy.gm2-subtitle-alt-1 > span:nth-child(1)"
route_string="#section-directions-trip-title-0 > span"

SampleNumber=5

trafficSheet.write(0,0,"Route")
trafficSheet.write(0,1,"Time spent")
trafficSheet.write(0,2,"Sampling date")
trafficSheet.write(0,3,"Sampling time")
i=1
driver=webdriver.Chrome()
driver.get(URL)

while(i<=SampleNumber):
    try: 
        driver.refresh()  
        time.sleep(5)
        time_spent=driver.find_element(By.CSS_SELECTOR,time_spent_string)
        route=driver.find_element(By.CSS_SELECTOR,route_string)
        now = datetime.now()
        time.sleep(5)
        time_spent_num=time_spent.text.split()[0]
        print("Date sampling number: "+str(i))
        print("Recommended route: "+route.text)
        print("Time spent: "+time_spent.text)       
        dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
        print("Sampling date and time: ", dt_string)
        trafficSheet.write(i,0,route.text)
        trafficSheet.write(i,1,int(time_spent_num))
        current_date=dt_string.split()[1]
        current_time=dt_string.split()[0]
        trafficSheet.write(i,2,current_time)
        trafficSheet.write(i,3,current_date)        
       
    except NoSuchElementException:
        print("Element not found")
        continue 
    print("***")
    i+=1
trafficWorkbook.close()
