from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import info

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(executable_path='chromedriver.exe', options=options)

url = 'https://www.bestbuy.com/site/nintendo-switch-oled-model-w-white-joy-con-white/6470923.p?skuId=6470923'
# url = 'https://www.bestbuy.com/site/philips-norelco-5300-wet-dry-electric-shaver-black-navy-blue/6384519.p?skuId=6384519'

driver.get(url)

isComplete = False

while not isComplete:
    try:
        addBtn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.add-to-cart-button'))
        )

    except:
        driver.refresh()
        continue

    print("Add to cart button found")

    try:
        addBtn.click()
        driver.get('https://www.bestbuy.com/cart')

        checkoutBtn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="cartApp"]/div[2]/div[1]/div/div[1]/div[1]/section[2]/div/div/div[3]/div/div[1]/button'))
        )
        checkoutBtn.click()
        print("Successfully added to cart - beginning check out")

        # fill in email and password
        emailField = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "fld-e"))
        )
        emailField.send_keys(info.email)

        pwField = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "fld-p1"))
        )
        pwField.send_keys(info.password)

        # click sign in button
        signInBtn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/section/main/div[2]/div[1]/div/div/div/div/form/div[3]/button'))
        )
        signInBtn.click()
        print("Signing in")

        isComplete = True

    except:
        driver.get(url)
        continue

print("Order successfully prepped")