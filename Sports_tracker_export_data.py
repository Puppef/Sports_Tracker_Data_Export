import numpy
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException  

import pandas as pd

import time
import datetime

#Connect to the Web Browser
browser = webdriver.Chrome(executable_path="C:/Users/Pierre/anaconda3/Lib/site-packages/selenium/chromedriver")
browser.set_window_position(-3000, 0) #move the tab away

df = pd.DataFrame() #Creation of the data frame   
#Export of the infos from Sports Tracker website to vectors and then to a data frame (df)
#date
def table():
    date = browser.find_elements_by_xpath("//span[@class='date']")
    date_list = [value.text for value in date]
    
#duration
    duration = browser.find_elements_by_xpath("//span[@class='duration']")
    duration_list = [value.text for value in duration]

#distance
    distance = browser.find_elements_by_xpath("//span[@class='distance']")
    distance_list = [value.text for value in distance]

#avg_speed
    avg_speed = browser.find_elements_by_xpath("//span[@class='avg-speed']")
    avg_speed_list = [value.text for value in avg_speed]

#energy
    energy = browser.find_elements_by_xpath("//span[@class='energy']")
    energy_list = [value.text for value in energy]
    
#Loop to insert infos into the df rows and columns 
    for i in range(len(date)):
        try:
            df.loc[i, 'date'] = date_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'duration'] = duration_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'distance'] = distance_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'avg-speed'] = avg_speed_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'energy'] = energy_list[i]
        except Exception as e:
            pass
    print('Table filled')
    
#Function to login into Sports Tracker
def enter_email(email, path_email, password, path_password):
    fly_from = browser.find_element_by_xpath(path_email)
    time.sleep(0.5)
    fly_from.clear()
    time.sleep(0.5)
    fly_from.send_keys(email)
    time.sleep(0.5)
    fly_from = browser.find_element_by_xpath(path_password)
    time.sleep(0.5)
    fly_from.clear()
    time.sleep(0.5)
    fly_from.send_keys(password)
    time.sleep(0.5)

#Function to press on the login button
def login():
    login = browser.find_element_by_xpath("//input[@class='submit']")
    login.click()
    time.sleep(2)
    print('Logged In')
    
#Main
link = 'https://www.sports-tracker.com/diary/workout-list/'
browser.get(link)
time.sleep(1)
enter_email('email@wtv.ocm', "//input[@id='username']", 'password' , "//input[@id='password']")
time.sleep(1)
login()
time.sleep(5)
browser.find_element_by_xpath("//select[@ng-model='activityType']/option[text()='Running']").click()  #click on running from the drop-down menu
time.sleep(2)
browser.find_element_by_xpath("//div[@class='show-more']").click()
time.sleep(1)
browser.find_element_by_xpath("//div[@class='show-more']").click()
time.sleep(2)
table()
time.sleep(1)
df.to_excel('Sprots_Tracker_Data.xlsx')   # file saved to "C:\Users\Pierre\.spyder-py3"
print('Excel generated!')