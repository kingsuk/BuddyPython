from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options


options = Options()
options.headless = False

browser = webdriver.Chrome(r"C:\Users\accenture.robotics\Downloads\chromedriver_win32 (2)\chromedriver",options=options)
browser.get("https://sanofiservices.service-now.com/nav_to.do?uri=%2Fx_saag_sanofi_its_incident_list.do%3Fsysparm_query%3Dassignment_group%253D33930e69db180bc005ea766eaf961951%255Eshort_descriptionSTARTSWITHUNX04%255EstateIN1%252C2%255Eassigned_toISEMPTY%26sysparm_first_row%3D1%26sysparm_view%3D")
# print(browser.current_window_handle)
# try:
#     alert = browser.switch_to_alert()
#     print(alert.text)
#     #alert.accept()
# except:
#     print("no alert to accept")

# print("url "+browser.current_url)



# searchBar = browser.find_element_by_name("q")
# searchBar.send_keys("Python And Selenium")
# searchBar.send_keys(Keys.ENTER)

# searchResults = browser.find_element_by_class_name("r")
# print("result = "+searchResults.text)

#browser.close()