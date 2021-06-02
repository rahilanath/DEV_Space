from selenium import webdriver

browser = webdriver.Chrome('C:\\Users\\Rahil\\Documents\\Github\\DEV_Space\\buybot\\chromedriver.exe')

browser.get('https://www.gamestop.com/video-games/nintendo-switch/consoles/products/nintendo-switch-animal-crossing-new-horizons-edition/212415.html?utm_source=google&utm_medium=feeds&utm_campaign=unpaid_listings')

addBtn = browser.find_element_by_class_name('add-to-cart')
addBtn.click()


added = False

while added == False:
    try:
        vwCart = browser.find_element_by_class_name('view-cart-button')
        vwCart.click()
        added = true

    except:
        pass