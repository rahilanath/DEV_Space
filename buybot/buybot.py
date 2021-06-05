from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])

browser = webdriver.Chrome(executable_path='chromedriver.exe', options=options)

# url = 'https://www.gamestop.com/video-games/nintendo-switch/consoles/products/nintendo-switch-animal-crossing-new-horizons-edition/212415.html?utm_source=google&utm_medium=feeds&utm_campaign=unpaid_listings'
url = 'https://www.gamestop.com'

browser.get(url)

wait = WebDriverWait(browser, 50)
wait.until(EC.presence_of_element_located((By.NAME, 'q')))

srchBox = browser.find_element_by_name('q')
srchBtn = browser.find_element_by_class_name('search-icon')

srchBox.send_keys('Nintendo Switch Pro')
srchBtn.click()

# addBtn = browser.find_element_by_class_name('add-to-cart')
# addBtn.click()

# added = False

# while added == False:
#     try:
#         vwCart = browser.find_element_by_class_name('view-cart-button')
#         vwCart.click()
#         added = true

#     except:
#         pass