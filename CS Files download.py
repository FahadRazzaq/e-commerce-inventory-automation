#!/usr/bin/env python
# coding: utf-8

# <center><h2> Channel Spyder

# #### Import Required Python Packages

# In[1]:


import pandas as pd
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
import os
import shutil
import time
from datetime import datetime
import zipfile
import slack_sdk
from datetime import date

#from webdriver_manager.chrome import ChromeDriverManager


# ### Paths

# In[2]:


inventories = os.getcwd() + '\inventories'
downloads = os.getcwd() + '\CS Inventories'


# Deleting the Downloads Folder and creating new

# In[3]:


if os.path.exists(downloads):
    shutil.rmtree(downloads)
    
os.mkdir(downloads)


# In[4]:


# Create ChromeOptions object to store login credentials
options = webdriver.ChromeOptions()

### Add argument to ChromeOptions object to specify location of user data for Chrome
udata_folder = os.getenv('LOCALAPPDATA') + r'\Google\Chrome\User Data\Channel Spyder'   #Path to Channel Spyder Selenium Profile
options.add_argument(f"--user-data-dir={udata_folder}")   #Using Selenium Chrome Profile to Open 

downloads = os.getcwd() + '\CS Inventories'   #Setting Path for the Downloads Folder
#Changing Default Download Path
prefs = {'download.default_directory' : downloads}
options.add_experimental_option('prefs', prefs)

## Create Chrome webdriver object, using ChromeOptions object and installing necessary driver with ChromeDriverManager
#browser = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))  #Opening selenium

browser = webdriver.Chrome(options=options)  #Opening selenium


# #### Target URL

# Login to Channel Spyder if not logged in

# In[5]:


# driver = webdriver.Chrome()
url = r"https://www.channelspyder.com/inventory/inventory_list?file_type=1&upload_type=0&limit=200" # paste your url here
browser.get(url)

if browser.current_url == 'https://www.channelspyder.com/login':
    browser.find_element(By.XPATH, '//input[@name="email_login"]').send_keys('billsmithauto')
    browser.find_element(By.XPATH, '//input[@name="password_login"]').send_keys('Kaag6202!')
    browser.find_element(By.XPATH, '//*[@id="loginForm"]/button[2]').click()
    browser.get(url)
    
    


# #### Download Files from Channel Spyder

# In[6]:


var = browser.find_elements(By.XPATH, '//*[@id="simpledatatable"]/tbody//tr/td[2]/a')


# In[7]:


# Create a dictionary called dictionary with keys being various warehouse names and values being their corresponding short names
dictionary = {'Turn14':'t14', 'Burco Mirrors':'burco', 'PFG':'pfg',
             'Keystone':'keystone', 'Parts Auth':'pa', 'LKQ':'lkq',
             'NPW':'npw', 'Wheel Pros':'wp', 'SimpleTire':'simpletire',
             'Race Sport Lighting':'rsl', 'Tonsa':'tonsa', 'Motor State':'motorstate',
             'OE Wheels':'oe', 'Jante Wheel':'jante', 'Dorman Direct':'dorman',
             }


# In[8]:


### Enter start and end index to download files on inventory page on channel spyder
start_index = 1
end_index  = 20


# In[ ]:


# Find all elements with a tag of "a"
var = browser.find_elements(By.XPATH, '//*[@id="simpledatatable"]/tbody//tr/td[2]/a')   #Getting all download buttons
var_dts = browser.find_elements(By.XPATH, '//*[@id="simpledatatable"]/tbody//tr/td[11]')   #Getting all Timestamp Strings
var_fnm = browser.find_elements(By.XPATH, '//*[@id="simpledatatable"]/tbody//tr/td[3]')   #Getting al Warehouse Names


## Store the path to inventory page in variable->Inventory_page
Inventory_page = browser.current_window_handle

#### Download files
# Loop through the range of numbers from start_index to end_index
for i in range(start_index, end_index):  
    # Try to click on the element at index i-1 of the var list
    # and print a message saying the file was downloaded
    # If the element at index i-1 does not exist, catch the exception and print an error message
    try:
        fname = var[i-1].get_attribute('title')   #Getting name of file which will be downloaded
        if 'inv_availability_allparts' in fname:   #Skip the other Dorman File
            continue
            
#       Downloading the file

        ## Get the URL of download button from Inventory page
        href = var[i-1].get_attribute('href')
        ## switches the browser to a new window
        browser.switch_to.new_window()  
        
        # Navigates the new window browser to the URL specified in the "href" variable.
        browser.get(href)
        print('Opening Page: ', i)
        time.sleep(1)
        
        try:
            browser.find_element(By.XPATH, '//div[@role="alert"]')   #Checking if the flie not found alert is shown
            print("File at index {} on Channel Spyder not found".format(i))
            # Close the new open window of the browser
            browser.close()  
            # switches the browser back to the original window-> Inventory page 
            browser.switch_to.window(Inventory_page)
            continue
        except:
            pass

            
        # Close the new open window of the browser
        browser.close()  
        
        # switches the browser back to the original window-> Inventory page 
        browser.switch_to.window(Inventory_page)
        
        
        t_out = 300   #Defining a timeout for safety
        while((not (os.path.exists(downloads + r'\\' + fname))) and (t_out>0)):   #Dynamically waiting for the download to finish
            time.sleep(1)
            t_out = t_out - 1
            
        d_str = var_dts[i-1].get_attribute('innerText')   #Getting Time Stamp String
        dttm = datetime.strptime(d_str, '%Y-%m-%d %H:%M:%S')     #Converting Time Stamp String to Python Timestamp Object
        
        wnm = var_fnm[i-1].get_attribute('innerText')   #Getting Warehouse name
        fnm = f"{dictionary[wnm]}_{dttm.strftime('%Y%m%d')}"   #Finalizing name of each file
        print (fnm)
        
        nm, ext = os.path.splitext(fname)   #Splitting the name and extension of the downloaded file
        if ext == '.zip':
            z = zipfile.ZipFile(f"{downloads}\\{nm}{ext}")
            zin = z.infolist()[0]
            zin.filename = f"{fnm}.csv"
            z.extract(member=zin, path=downloads)
            z.close()
            del z
            os.remove(f"{downloads}\\{nm}{ext}")
            
            
        else:
            os.rename(f"{downloads}\\{nm}{ext}", f"{downloads}\\{fnm}{ext}")
        
        print("Downloaded file no:", i)
    except:
        print("File at index {} on Channel Spyder not found".format(i))
    
browser.close()   #Close Chrome


# Coping all Files from Downloads to the inventories folder

# In[10]:


shutil.copytree(downloads, inventories, dirs_exist_ok=True)


# #### Copy files from downloads folder and save renamed files in Daily comparison folder
# All files are copy, and pasted with newnames except for Part Authority file that is downloaded in zip. We read file in pandas with zip compression and then save into csv 

# ### Zip all Inventory Downloaded files

# In[11]:


zipfilepath = 'inventories_' + datetime.now().strftime("%Y%m%d_%H%M%S")


# In[12]:


shutil.make_archive(zipfilepath, format='zip', root_dir=downloads)


# In[13]:


# import the slack_sdk module
import slack_sdk
from datetime import date

current_date = date.today()

# store the channel id for the slack channel where the message and file will be sent
cid = 'C04D1HF1RHT'

# create a client for the slack API using the WebClient class and the specified token
c = slack_sdk.WebClient(token='xoxb-156915382752-4437979958069-2XlrUKvulSUOltDeqdlHrqjA')

# create the text for the message to be sent, using the current date
text = 'Inventory Files ' + current_date.strftime("%d-%b-%Y")

# send the image file "df_styled.png" to the specified slack channel
c.files_upload_v2(file= f'{zipfilepath}.zip',  channel=cid, initial_comment=text)


# In[ ]:





# <center><h3> Files saved to Daily Comparison folder
