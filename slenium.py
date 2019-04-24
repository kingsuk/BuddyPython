from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

options = Options()
options.headless = False

browser = webdriver.Chrome(r"C:\Users\accenture.robotics\Downloads\chromedriver_win32\chromedriver",chrome_options=options)
browser.get("https://www.google.com/")
searchBar = browser.find_element_by_name("q")
searchBar.send_keys("Python And Selenium")
searchBar.send_keys(Keys.ENTER)

searchResults = browser.find_element_by_class_name("r")
print("result = "+searchResults.text)

browser.close()