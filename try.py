from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Setup the driver with options to disable pop-ups and notifications
edge_options = webdriver.EdgeOptions()
edge_options.add_argument("-private")
edge_options.add_argument("--disable-popup-blocking")
edge_options.add_argument("--disable-notifications")
edge_options.add_argument("--disable-infobars")
edge_options.add_argument("--disable-extensions")
edge_options.add_argument("--disable-blink-features=AutomationControlled")

# Specify the path to the WebDriver executable using the Service class
service = Service(executable_path='/path/to/msedgedriver')

driver = webdriver.Edge(service=service, options=edge_options)

# Open a website
driver.get('https://www.example.com')

# Check for private mode by looking for the private browsing indicator
private_indicator = driver.find_elements(By.CSS_SELECTOR, '.private-browsing-indicator')
if private_indicator:
    print("Private mode is enabled.")
else:
    print("Private mode is not enabled.")

# Check for pop-ups and notifications
try:
    WebDriverWait(driver, 10).until(EC.alert_is_present())
    print("Pop-up detected.")
except:
    print("No pop-ups detected.")

# Check for infobars
infobar = driver.find_elements(By.CSS_SELECTOR, '.infobar')
if infobar:
    print("Infobar is present.")
else:
    print("Infobar is not present.")

driver.quit()
