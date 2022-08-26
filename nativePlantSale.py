from argparse import Action
from tkinter import E
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from bs4 import BeautifulSoup
import requests
import os
from time import sleep


###Set working directory, file name
os.chdir(r"C:\Users\gabee\Desktop\Toledo Native Plant Sale")
fileName = "nativePlantsToledo.xlsx"
###Create blank dataframe
columns = ["name", "3.25_avaliable", "1_gallon_avaliable", "plant_type", "light_requirement", "soil_moisture", "height_in_ft", "bloom_time", "bloom_color", "attracts"]
df = pd.DataFrame(columns = columns)

###Open browser
url = "https://wildtoledo.org/collections/wild-toledo-native-plants-sales?view=24"
path = r"C:\Program Files (x86)\Geckodriver"
driver = webdriver.Firefox(path)
driver.get(url)
driver.maximize_window()


###Select load more button until all items are loaded
while True:
    try:
        loadMoreButton = driver.find_element(By.XPATH, "/html/body/main/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div/div/a/span")
        loadMoreButton.click()
        sleep(1)
    except Exception as e:
        print(e)
        break


###scroll back up to the top
driver.execute_script("window.scrollTo(0, 0)")
sleep(1.5)

###Get all of the items
elems = driver.find_elements(By.CSS_SELECTOR, "h4.h6.m-0")


###Open each item in a new tab and scrape the information
driver.switch_to.window(driver.window_handles[-1])


for i in range(len(elems)):
    #reset values
    name = ""
    pot_325_avaliable = ""
    pot_1_gallon_avaliable = ""
    plant_type = ""
    light_requirement = ""
    soil_moisture = ""
    height_in_ft = ""
    bloom_time = ""
    bloom_color = ""
    attracts = ""
    
    
    found=False
    while found==False:
        try:
            while(len(driver.window_handles)<2):
                ActionChains(driver).key_down(Keys.CONTROL).click(elems[i]).key_up(Keys.CONTROL).perform()
                if len(driver.window_handles)<2:
                    yScroll = int(driver.execute_script("return window.scrollY"))+(533/4)
                    xScroll = int(driver.execute_script("return window.scrollX"))
                    driver.execute_script("window.scrollTo({0}, {1})".format(xScroll, yScroll))
                if len(driver.window_handles)>1:
                    found = True
        except: #scroll down farther if the item isn't in view
            yScroll = int(driver.execute_script("return window.scrollY"))+(533/2)
            xScroll = int(driver.execute_script("return window.scrollX"))
            driver.execute_script("window.scrollTo({0}, {1})".format(xScroll, yScroll))
    
    
    ##In product page, pull the information
    driver.switch_to.window(driver.window_handles[-1])
    sleep(5)
    
    #get product name
    nameFound = False
    try:
        name = driver.find_element(By.XPATH,"/html/body/main/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/div[2]/h1")
        name = name.get_attribute("textContent")
    except:
        try:#Obedient plant didn't like the selector, but this one works.
            name = driver.find_element(By.CSS_SELECTOR, ".h3")
            name = name.get_attribute("textContent")
            nameFound = True
        except:
            print("name not found")
            name = "NA"
        if nameFound == False:
            print("name not found")
            name = "NA"
    
    #get sizes avaliable
        #3.25 inch pot
    try:
        pot_325 = driver.find_element(By.XPATH, "/html/body/main/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/div[7]/form/div[1]/div/div/div/div[1]")
        pot_325_class = pot_325.get_attribute("class")
        if "disabled" in pot_325_class:
            pot_325_avaliable = False
        if "active" in pot_325_class:
            pot_325_avaliable = True
    except:
        print("no 3.25 inches avaliable")
        pot_325_avaliable = False
    
        #1 gallon pot
    try:
        pot_1_gallon = driver.find_element(By.XPATH, "/html/body/main/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/div[7]/form/div[1]/div/div/div/div[2]")
        pot_1_gallon_class = pot_1_gallon.get_attribute("class")
        if "disabled" in pot_1_gallon_class:
            pot_1_gallon_avaliable = False
        if "active" in pot_1_gallon_class:
            pot_1_gallon_avaliable = True
    except:
        print("no 1 gallon pots avaliable")
        pot_1_gallon_avaliable = False
    
    try:
        url = driver.current_url
        content = requests.get(url).text
        soup = BeautifulSoup(content, 'html.parser') 
        pageRaw = soup.find_all("div", class_="tabs__content rte overflow-hidden")
        page = str(pageRaw)
        page = page.replace("\xa0", " ")
        page = page.replace("&amp;", "&")
        
        #Find the scientific name
        beginning = page.find("<em>")+4
        end = page[beginning:].find("</em>")
        sciName = page[beginning:beginning+end]
        if page.find("<em>")==-1:
            sciName = "NA"
            
        #Find plant type
        beginning = page.find("Plant type: ")+len("Plant type: ")
        end = page[beginning:].find("</p>")
        plant_type = page[beginning:beginning+end]
        if page.find("Plant type: ")==-1:
            plant_type = "NA"
            
        #Find light requirement
        beginning = page.find("Light requirement: ")+len("Light requirement: ")
        end = page[beginning:].find("</p>")
        light_requirement = page[beginning:beginning+end]
        if page.find("Light requirement: ")==-1:
            light_requirement = "NA"
            
        #Find moisture
        beginning = page.find("Soil moisture: ")+len("Soil moisture: ")
        end = page[beginning:].find("</p>")
        soil_moisture = page[beginning:beginning+end]
        if page.find("Soil moisture: ")==-1:
            soil_moisture = "NA"
            
        #find height
        beginning = page.find("Height (in feet): ")+len("Height (in feet): ")
        end = page[beginning:].find("</p>")
        height_in_ft = page[beginning:beginning+end]
        if page.find("Height (in feet): ")==-1:
            height_in_ft = "NA"
            
        #find bloom time
        beginning = page.find("Bloom time: ")+len("Bloom time: ")
        end = page[beginning:].find("</p>")
        bloom_time = page[beginning:beginning+end]
        if page.find("Bloom time: ")==-1:
            bloom_time = "NA"
            
        #find bloom color
        beginning = page.find("Bloom color: ")+len("Bloom color: ")
        end = page[beginning:].find("</p>")
        bloom_color = page[beginning:beginning+end]
        if page.find("Bloom color: ")==-1:
            bloom_color = "NA"
            
        #find attracts
        beginning = page.find("Attracts:")+len("Attracts:")
        end = page[beginning:].find("</p>")
        attracts = page[beginning:beginning+end]
        if page.find("Attracts:")==-1:
            attracts = "NA"
    except:
        print("something went wrong")
        
    #insert into row
    try:
        df.loc[len(df)] = [name, pot_325_avaliable, pot_1_gallon_avaliable, plant_type, light_requirement, soil_moisture, height_in_ft, bloom_time, bloom_color, attracts]
    except:
        print("unable to insert record " + i)
        
    
    #close webpage and go back to main page
    driver.close()
    driver.switch_to.window(driver.window_handles[-1])
    sleep(1)


print("finished extracting data")
driver.close()

###Write dataframe to sheet
df.to_excel(fileName)

####Testing
"""writer = pd.ExcelWriter(fileName, engine='xlsxwriter')
workbook = writer.book
worksheet = writer.sheets['Sheet1']"""