# -*- coding: utf-8 -*-
"""
Created on Fri Dec 24 11:19:24 2021

@author: duxin
"""
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException

import xlsxwriter
outWorkbook=xlsxwriter.Workbook("Home_to_Florica_Chemical_Co.xlsx")
outSheet=outWorkbook.add_worksheet()
URL="https://www.google.com/maps/dir/14563+Global+Cir+%236202,+Orlando,+FL+32821/351+Bert+Schulz+Boulevard,+Winter+Haven,+FL/@28.20192,-81.762894,11z/data=!3m1!4b1!4m14!4m13!1m5!1m1!1s0x88dd81cf2130f55d:0xe10322359e81f3da!2m2!1d-81.5004462!2d28.3522537!1m5!1m1!1s0x88dd6d6809c9e145:0x67cbe9012e8ccfb6!2m2!1d-81.7021904!2d28.0727179!3e0"

# outSheet.write("A1","Route")
# outSheet.write("B1","Time spent")
# outSheet.write("C1","Sampling date and time")

outSheet.write(0,0,"Route")
outSheet.write(0,1,"Time spent")
outSheet.write(0,2,"Sampling date")
outSheet.write(0,3,"Sampling time")

driver=webdriver.Chrome()
#driver.get("https://www.google.com/maps/dir/14563+Global+Cir+%236202,+Orlando,+FL+32821/4304+Scorpius+Street,+Orlando,+FL/@28.4771767,-81.4935348,11z/data=!3m1!4b1!4m14!4m13!1m5!1m1!1s0x88dd81cf2130f55d:0xe10322359e81f3da!2m2!1d-81.5004462!2d28.3522537!1m5!1m1!1s0x88e7685cc9b7552b:0x100cb38b5c8b277a!2m2!1d-81.1971585!2d28.601259!3e0") 
driver.get(URL)
print(driver.title)
#driver.get("https://www.bridgewebs.com/orlandometrobridgecenter/")
#text=driver.find_elements_by_xpath("//p[contains(text(),'fastest route']")
#text=driver.find_element_by_xpath("//*[@id="section-directions-trip-0"]/div/div[1]/div[3]/span[1]/span[1]/span[2]")
#text=driver.find_element_by_class_name("LLgQof-LYNcwc-text renderable-component-text-not-line")
#text=driver.find_element_by_css_selector("#section-directions-trip-0 > div > div:nth-child(1) > div.xB1mrd-T3iPGc-iSfDt-HSrbLb.xB1mrd-T3iPGc-iSfDt-K4efff-text.gm2-body-2 > span.renderable-component > span:nth-child(2) > span.LLgQof-LYNcwc-text.renderable-component-text-not-line").text
i=1
while(i<2000):
    try: 
        #driver.get("https://www.google.com/maps/dir/14563+Global+Cir+%236202,+Orlando,+FL+32821/4304+Scorpius+Street,+Orlando,+FL/@28.4771767,-81.4935348,11z/data=!3m1!4b1!4m14!4m13!1m5!1m1!1s0x88dd81cf2130f55d:0xe10322359e81f3da!2m2!1d-81.5004462!2d28.3522537!1m5!1m1!1s0x88e7685cc9b7552b:0x100cb38b5c8b277a!2m2!1d-81.1971585!2d28.601259!3e0") 
        driver.get(URL)
        time.sleep(5)
        time_spent=driver.find_element_by_css_selector("#section-directions-trip-0 > div > div:nth-child(1) > div.xB1mrd-T3iPGc-iSfDt-n5AaSd > div.xB1mrd-T3iPGc-iSfDt-duration.delay-light.gm2-subtitle-alt-1 > span:nth-child(1)")
        route=driver.find_element_by_css_selector("#section-directions-trip-title-0 > span")
        now = datetime.now()
        time.sleep(5)
        time_spent_num=time_spent.text.split()[0]
        print(time_spent.text)
        print(route.text)
        dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
        print("sampling time: ", dt_string)
        outSheet.write(i,0,route.text)
        outSheet.write(i,1,int(time_spent_num))
        current_date=dt_string.split()[0]
        current_time=dt_string.split()[1]
        outSheet.write(i,2,current_time)
        outSheet.write(i,3,current_date)
        
       
    except NoSuchElementException:
        print("Element not found")
        continue
    print(i)
    i+=1
outWorkbook.close()
#search= driver.find_elements_by_name("s")
#search[0].send_keys("test")
#search[0].send_keys(Keys.RETURN)
#print(driver.page_source)
#time.sleep(6)
# try:
#     main=WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"main")))
#     print(main.text)
    
#     # articles = main.find_elements_by_tag_name("article")
#     # for article in articles:
#     #     header=article.find_element_by_class_name("entry_summary")
#     #     print(header.text)
# except:
#     driver.quit()
# main=driver.find_element_by_id("main")
# print(main.text)

# link=driver.find_element_by_link_text("Python Programming")
# #time.sleep(5)
# link.click()

# try:
#     element=WebDriverWait(driver,10).until(EC.presence_of_element_located((By.LINK_TEXT,"Beginner Python Tutorials")))
#     #print(main.text)
#     element.click()
#     # articles = main.find_elements_by_tag_name("article")
#     # for article in articles:
#     #     header=article.find_element_by_class_name("entry_summary")
#     #     print(header.text)
# except:
#     driver.quit()
    
# driver.back()
# driver.back()
# driver.forward()
#<span jstcache="204" class="LLgQof-LYNcwc-text renderable-component-text-not-line" jsan="7.LLgQof-LYNcwc-text,7.renderable-component-text-not-line">F