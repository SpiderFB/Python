from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random

# Setup the driver with options to disable pop-ups and notifications
edge_options = webdriver.EdgeOptions()
edge_options.add_argument("-private")
# options.add_argument("--incognito") #For Chrome
edge_options.add_argument("--disable-popup-blocking")
edge_options.add_argument("--disable-notifications")
edge_options.add_argument("--disable-infobars")
edge_options.add_argument("--disable-extensions")
edge_options.add_argument("--disable-blink-features=AutomationControlled")
# edge_options.add_argument("--headless") #it will operate without opening a visible browser window

driver = webdriver.Edge(options=edge_options)
count = 0

driver.get('https://vdi.portal.effem.com/Citrix/AP-STRWeb/')
while True:
    try:
        input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'input28'))
        )
        # Fill the input field with the desired text
        input_field.send_keys('arnab.kumar.roy@effem.com')
        
        NEXT_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form20"]/div[2]/input')))
        NEXT_btn.click()

        
        Detect_Citrix_Workspace_app = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="protocolhandler-welcome-installButton"]'))
        )
        Detect_Citrix_Workspace_app.click()




        # # Handle the pop-up for Citrix Workspace Launcher
        # open_button = WebDriverWait(driver, 10).until(
        #     EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Open")]'))
        # )
        # open_button.click()




        Already_installed = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="protocolhandler-detect-alreadyInstalledLink"]'))
        )
        Already_installed.click()



        wait = WebDriverWait(driver, 10)
        checkbox = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='checkbox' and @name='alwaysAllow']")))





        # HOME_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="myHomeBtn"]/div')))
        # HOME_btn.click()
        
        MPWAP_drop_down = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="home-screen"]/div[2]/section[5]/div[5]/div/ul/li[1]/a[2]/img')))
        MPWAP_drop_down.click()

    except Exception as e:
        print(f"Attempt {count}: {e}")
        count += 1
        time.sleep(random.uniform(1, 3))



