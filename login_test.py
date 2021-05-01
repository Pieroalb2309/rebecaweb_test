import time
import os
import excelfunctions as xlc
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

path = os.path.abspath("datatest.xlsx")
driver = xlc.start_webdriver()
driver.get("http://localhost:3000/auth/login")
driver.maximize_window()
print(driver.title)
rows=xlc.getrowcount(path,'login')
#LOG IN---
for r in range(2,rows+1):
    username=xlc.readdata(path,'login',r,1)
    password=xlc.readdata(path,'login',r,2)
    driver.find_element_by_id("login-form_email").send_keys(username)
    driver.find_element_by_id("login-form_password").send_keys(password)
    driver.find_element_by_xpath("//button[@type='submit']").click()
    time.sleep(3)
    #print(driver.find_element_by_id('myteam').is_displayed())
    try:
        if (driver.find_element_by_id('myteam').is_displayed() == True):
            print("success")
            xlc.writedata(path, 'login', r, 3, 'success')
            xlc.writedata(path, 'login', r, 4, "")
        driver.find_element_by_id("navprofile").click()
        time.sleep(1)
        driver.find_element_by_xpath("//span[text()='Cerrar sesión']").click()
        pass
    except:
        driver.find_element_by_id("login-form_email").send_keys(Keys.CONTROL + "a")
        driver.find_element_by_id("login-form_email").send_keys(Keys.DELETE)
        driver.find_element_by_id("login-form_password").send_keys(Keys.CONTROL+"a")
        driver.find_element_by_id("login-form_password").send_keys(Keys.DELETE)
        print('failed')
        msg = driver.find_element_by_id("alert_msg").text
        xlc.writedata(path, 'login', r, 3, 'failed')
        xlc.writedata(path, 'login', r, 4, msg)

        pass

print('finish')
driver.close()








