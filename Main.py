from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import  NoSuchElementException
import pandas as pd
import os
import time
import random

# Product = ["423149702","423149700","423149683","423149677","423149667"]
Product = ["423149702"]


user_name = os.getlogin()

# def meesho_automation(name, contact_Number, house_no, pincode, city, state):

options = webdriver.ChromeOptions()
# add your chrome driver path
print(user_name)
options.add_argument(f"--user-data-dir=C:\\Users\\{user_name}\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument(r'--profile-directory=Profile 4')
driver_path = r"chromedriver-win64\chromedriver.exe"
service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=options)

# Open Meesho website
driver.get("https://www.meesho.com/")
time.sleep(random.uniform(3, 5))  # Random delay

def meesho_automation(name, contact_Number, house_no, pincode, city, state):

    for product in Product:
        search_input = driver.find_element(By.XPATH, "//input[@placeholder='Try Saree, Kurti or Search by Product Code']")
        # search_input.send_keys(product_name)
        search_input.clear()
        for digit in product:
            search_input.send_keys(digit)
            time.sleep(random.uniform(0.1, 0.3))
        search_input.send_keys(Keys.RETURN)
        time.sleep(5)
        # Wait for search results and click on the first product
        first_product = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div[class*='ProductList__GridCol']"))
        )
        first_product.click()

        add_to_cart_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Add to Cart']"))
        )
        add_to_cart_button.click()

    cart_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Cart']"))
    )
    cart_button.click()

    buy_now_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Continue']"))
    )
    buy_now_button.click()
    
    ADD_NEW_ADDRESS = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='ADD NEW ADDRESS']"))
    )
    ADD_NEW_ADDRESS.click()

    time.sleep(1)

    Name = driver.find_element(By.XPATH, "//input[@id='name']")
    Name.send_keys(name)
    Name.send_keys(Keys.RETURN)
    time.sleep(1)
    Contact_Number = driver.find_element(By.XPATH, "//input[@id='mobile']")
    Contact_Number.send_keys(contact_Number)
    Contact_Number.send_keys(Keys.RETURN)
    
    House_no = driver.find_element(By.XPATH, "//input[@id='addressLine1']")
    House_no.send_keys(house_no)
    House_no.send_keys(Keys.RETURN)

    House_no_1 = driver.find_element(By.XPATH, "//input[@id='addressLine2']")
    House_no_1.send_keys(house_no)
    House_no_1.send_keys(Keys.RETURN)

    Pincode = driver.find_element(By.XPATH, "//input[@id='pincode']")
    Pincode.send_keys(pincode)
    Pincode.send_keys(Keys.RETURN)

    City = driver.find_element(By.XPATH, "//input[@id='city']")
    City.send_keys(city)
    City.send_keys(Keys.RETURN)



    # Wait for the dropdown options to appear
    state_input = driver.find_element(By.ID, "state")  # Adjust the selector as needed

    # Click on the state input to open the dropdown
    state_input.click()

    dropdown = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.sc-iBPTVF.gDPWIS"))
    )

    # Function to find and click a state
    def select_state(state_name):
        state_elements = dropdown.find_elements(By.CSS_SELECTOR, "div.sc-iBPTVF.ddDdMx")
        for element in state_elements:
            if element.text.strip() == state_name:
                element.click()
                return True
        return False

    desired_state = state
    if select_state(desired_state):
        print(f"Successfully selected {desired_state}")
    else:
        print(f"Failed to select {desired_state}")

    SAVE_ADD_NEW_ADDRESS = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Save Address and Continue']"))
    )
    SAVE_ADD_NEW_ADDRESS.click()

    time.sleep(2)
    
    address_containers = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.sc-iBPTVF.bRcHgY"))
    )

    address_selected = False
    for container in address_containers:
        # If a name is provided, check if it matches
        if name:
            try:
                name_element = container.find_element(By.CSS_SELECTOR, "h6.sc-bdfCDU.iGbpxC")
                if name_element.text.strip() == name:
                    container.click()
                    print(f"Selected address for {name}")
                    address_selected = True
                    break
            except NoSuchElementException:
                continue
        else:
            # If no name is provided, select the first address
            container.click()
            print("Selected the first available address")
            address_selected = True
            break

    if address_selected:
        # Attempt to confirm the address selection
        try:
            confirm_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Deliver to this Address']"))
            )
            confirm_button.click()
            print("Confirmed address selection")
        except:
            print("Could not find or click the confirmation button")

    Payment_Method = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Continue']"))
    )
    Payment_Method.click()

    Place_Order = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Place Order']"))
    )
    Place_Order.click()

    time.sleep(1)

    Place_Order = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Continue Shopping']"))
    )
    Place_Order.click()



    time.sleep(1) 


# Read the Excel file
df = pd.read_excel('200 od.xlsx')

# Usage
for index, row in df.iterrows():
    meesho_automation(
        row['Customer Name'],
        row['Mobile Number'],
        row['Address'],
        row['Pincode'],
        row['City'],
        row['State']
    )
