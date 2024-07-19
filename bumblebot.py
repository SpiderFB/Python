from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random

# Setup the driver. This one uses Edge, but you can use others too.
driver = webdriver.Edge()
count = 0

# Connect to the webpage
driver.get('https://gew3.bumble.com/app')

while True:
    try:
        # Find the button and click it
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/div/div[1]/main/div[2]/div/div/span/div[2]/div/div[2]/div/div[3]/div/div[1]/span')))
        button.click()
        count += 1
        print ("Right Swiped----------------------->" + str(count))
        time.sleep(random.randint(5, 25))
    except Exception as e:
        # If the button is not found because the page hasn't loaded yet, wait a bit and try again
        print(f"Exception occurred: {str(e)}")
        time.sleep(random.randint(5, 25))
