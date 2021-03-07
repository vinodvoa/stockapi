import os
from selenium import webdriver

if os.path.exists('/Users/vinodverghese/Downloads/finviz.csv'):
    os.remove('/Users/vinodverghese/Downloads/finviz.csv')

profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)  # custom location
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', '/Users/vinodverghese/Downloads')
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
profile.set_preference("browser.link.open_newwindow", 3)
profile.set_preference("browser.link.open_newwindow.restriction", 2)

driver = webdriver.Firefox(firefox_profile=profile)

driver.get('https://finviz.com/groups.ashx?g=industry&v=120&o=pe')

export = driver.find_element_by_xpath('/html/body/table[3]/tbody/tr[5]/td/table/tbody/tr/td/a')
export.click()

driver.quit()
