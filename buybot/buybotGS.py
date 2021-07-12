from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from notify_run import Notify
# notify-run configure https://notify.run/OJvC9FrZQAhSfZ7b

import info

notify = Notify()

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(executable_path='chromedriver.exe', options=options)

switch_oled_url = 'https://www.gamestop.com/video-games/nintendo-switch/consoles/products/nintendo-switch-oled-with-white-joy-con/11149258.html?condition=New'
ps5_url = 'https://www.gamestop.com/video-games/playstation-5/consoles/products/playstation-5-digital-edition/11108141.html?condition=New'
test_url = 'https://www.gamestop.com/video-games/nintendo-switch/consoles/products/nintendo-switch-with-neon-blue-and-neon-red-joy-con/11095819.html?condition=New'

# SET WHICH PRODUCT TO SNIPE
url = test_url

driver.get(url)

isComplete = False

while not isComplete:
    try:
        addBtn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.add-to-cart'))
        )

    except:
        driver.refresh()
        continue

    notify.send('ADDED ITEM TO CART HURRY')

    try:
        addBtn.click()
        driver.get('https://www.gamestop.com/cart/')

        isComplete = True

    except:
        driver.get(url)
        continue

print("Added to cart...")