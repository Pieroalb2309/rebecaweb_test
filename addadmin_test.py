import time
import os
import excelfunctions as xlc
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

path = os.path.abspath("datatest.xlsx")
driver = xlc.start_webdriver()

driver.get("http://localhost:3000/auth/login")
driver.maximize_window()
rows=xlc.getrowcount(path,'admin-create')
print(driver.title)
#Test date---
today = date.today()
d1 = today.strftime("%d/%m/%Y")
#LOG IN---
msg="not logged"
driver.find_element_by_id("login-form_email").send_keys('sergio.sotelo@eulerinnovations.com')
driver.find_element_by_id("login-form_password").send_keys('eulerinnovations')
driver.find_element_by_xpath("//button[@type='submit']").click()
time.sleep(5)
print("logged!!")
#inicia el bucle
for r in range(3,rows+1):
    try:
        msg = "FORMULARIO NUMERO: " + str(r - 2)  #####################################################
        print(msg)
        flag = 0
        nombre = xlc.readdata(path, 'admin-create', r, 1)
        apellido = xlc.readdata(path, 'admin-create', r, 2)
        permiso = xlc.readdata(path, 'admin-create', r, 3)
        proveedor = xlc.readdata(path, 'admin-create', r, 4)
        correo = xlc.readdata(path, 'admin-create', r, 5)
        password = xlc.readdata(path, 'admin-create', r, 6)
        confirm_pass = xlc.readdata(path, 'admin-create', r, 7)
        mobile = xlc.readdata(path, 'admin-create', r, 8)
        # ingresar al formulario
        #driver.find_element_by_id("rc-tabs-0-tab-2").click()
        driver.find_element_by_xpath("//div[text()='Usuarios']").click()
        time.sleep(3)
        driver.find_element_by_xpath("(//button[text()=' Agregar'])[2]").click()
        time.sleep(3)
        print("ingreso al formulario")
        msg = "Inform. Basica"
        print(msg)
        time.sleep(2)
        driver.find_element_by_id("register_form_first_name").click()
        time.sleep(1)
        driver.find_element_by_id("register_form_first_name").send_keys(nombre)
        time.sleep(1)
        driver.find_element_by_id("register_form_last_name").click()
        time.sleep(1)
        driver.find_element_by_id("register_form_last_name").send_keys(apellido)
        time.sleep(1)
        msg = "Permiso"
        print(msg)
        group = xlc.permissions(permiso)
        for i in range(len(group)):
            driver.find_element_by_id("register_form_permission_groups").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[@title='"+group[i]+"']").click()
            time.sleep(1)
        msg = "Proveedor"
        print(msg)
        #driver.find_element_by_id("register-form_provider").send_keys(proveedor)
        driver.find_element_by_xpath("(//input[@class='ant-select-selection-search-input'])[2]").send_keys(proveedor)
        time.sleep(1)
        driver.find_element_by_xpath("//div[text()='"+proveedor+"']").click()
        time.sleep(1)
        driver.find_element_by_id("register_form_email").send_keys(correo)
        time.sleep(1)
        driver.find_element_by_id("register_form_password").send_keys(password)
        time.sleep(1)
        driver.find_element_by_id("register_form_conf_password").send_keys(confirm_pass)
        time.sleep(1)
        driver.find_element_by_id("register_form_mobile").send_keys(mobile)
        time.sleep(1)

        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
        time.sleep(3)
        print("Subiendo....")
        driver.find_element_by_xpath("//button[text()='Enviar']").click()
        flag=1
        time.sleep(15)
        if (driver.find_element_by_id("myteam").is_displayed() == True):
            msg = "FORMULARIO LLENADO NUMERO: " + str(r - 2)  #####################################################
            print(msg)
            xlc.writedata(path, "admin-create", r, 9, 'success')
            xlc.writedata(path, "admin-create", r, 10, '')
            xlc.writedata(path, "admin-create", r, 11, d1)
        #driver.find_element_by_xpath("(//button[@type='button'])[1]").click()

        pass
    except:
        #value = driver.find_element_by_xpath("(//div[@role='alert'])[1]").get_attribute('innerHTML').strip()
        #value = value[683:]
        #value = value[:(len(value) - 53)]
        # xlc.writedata(path, 'login', r, 3, msg)
        print('failed')
        xlc.writedata(path, "admin-create", r, 11, d1)
        xlc.writedata(path, "admin-create", r, 9, 'failed')
        if flag == 0:
            xlc.writedata(path, "admin-create", r, 10, msg)
        else:
            xlc.writedata(path, "admin-create", r, 10, 'Backend Time limit reached')
        # 3 segundos para subir al inicio del formulario y click en el boton de regreso
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
        time.sleep(3)
        driver.find_element_by_xpath("(//button[@type='button'])[1]").click()
        pass

print("end")
driver.close()