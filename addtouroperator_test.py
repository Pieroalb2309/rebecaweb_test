import time
import os
import excelfunctions as xlc
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

path = os.path.abspath("datatest.xlsx")
driver = xlc.start_webdriver()

driver.get("http://localhost:3000/auth/login")
driver.maximize_window()
rows=xlc.getrowcount(path,'tourop-create')
print(driver.title)
#Create tour operator---
msg="not created"
for r in range(3,rows+1):
    driver.find_element_by_xpath("//a[@href='/auth/register-3']").click()
    time.sleep(1)
    try:
        msg = "FORMULARIO NUMERO: " + str(r - 2)  #####################################################
        print(msg)
        flag=0
        company_name=xlc.readdata(path, 'tourop-create', r, 1)
        company_mobile=xlc.readdata(path, 'tourop-create', r, 2)
        descripcion=xlc.readdata(path, 'tourop-create', r, 3)
        local_pais=xlc.readdata(path, 'tourop-create', r, 4)
        local_estado=xlc.readdata(path, 'tourop-create', r, 5)
        local_ciudad=xlc.readdata(path, 'tourop-create', r, 6)
        cod_postal=xlc.readdata(path, 'tourop-create', r, 7)
        direccion=xlc.readdata(path, 'tourop-create', r, 8)
        cobertura_pais=xlc.readdata(path, 'tourop-create', r, 9)
        cobertuda_estado=xlc.readdata(path, 'tourop-create', r, 10)
        contact_nombre=xlc.readdata(path, 'tourop-create', r, 11)
        contact_apellido=xlc.readdata(path, 'tourop-create', r, 12)
        contact_correo=xlc.readdata(path, 'tourop-create', r, 13)
        contact_mobile=xlc.readdata(path, 'tourop-create', r, 14)
        mainuser_nombre=xlc.readdata(path, 'tourop-create', r, 15)
        mainuser_apellido=xlc.readdata(path, 'tourop-create', r, 16)
        mainuser_mobile=xlc.readdata(path, 'tourop-create', r, 17)
        mainuser_email=xlc.readdata(path, 'tourop-create', r, 18)
        mainsuser_pass=xlc.readdata(path, 'tourop-create', r, 19)
        mainuser_confpass=xlc.readdata(path, 'tourop-create', r, 20)
        msg = "basic info"#####################################################
        print(msg)
        driver.find_element_by_id("register-form_company_name").send_keys(company_name)
        driver.find_element_by_id("register-form_company_mobile").send_keys(company_mobile)
        driver.find_element_by_id("register-form_description-es").send_keys(descripcion)
        time.sleep(1)
        #driver.find_element_by_xpath("//button[text()='Traducción']").click()
        time.sleep(5)
        driver.find_element_by_xpath("(//span[text()='Ingles'])[2]").click()
        msg = "inform. local"#####################################################
        print(msg)
        driver.find_element_by_id("register-form_country").send_keys(local_pais)
        time.sleep(1)
        driver.find_element_by_xpath("//div[text()='"+local_pais+"']").click()#(//div[@class='ant-select-item-option-content'])[1]
        # state
        msg = "Local State"
        print(msg)
        driver.find_element_by_id("register-form_state").send_keys(local_estado)#//div[text()='"+local_estado+"']
        time.sleep(1)
        value = (driver.find_element_by_xpath("(//div[@class='ant-select-item-option-content'])[2]").get_attribute('innerHTML').strip())
        if value[0] == local_estado[0]:
            driver.find_element_by_xpath("(//div[@class='ant-select-item-option-content'])[2]").click()
        elif local_estado in value:
            driver.find_element_by_xpath("//div[text()='" + value + "']").click()
        else:
            driver.find_element_by_xpath("//div[text()='"+local_estado+"']").click()
        # city
        #msg = "Local City"
        #print(msg)
        #if local_ciudad != "":
        #    driver.find_element_by_id("register-form_city").send_keys(local_ciudad)#//div[text()='"+local_ciudad+"']
        #    driver.find_element_by_xpath("//div[text()='"+local_ciudad+"']").click()
        driver.find_element_by_id("register-form_zipcode").send_keys(cod_postal)
        driver.find_element_by_id("register-form_address").send_keys(direccion)
        msg = "zona cobertura"#####################################################
        print(msg)
        driver.find_element_by_id("register-form_coverage_0_country").send_keys(cobertura_pais)#//div[text()='"+cobertura_pais+"']
        time.sleep(1)
        if cobertura_pais == local_pais:
            driver.find_element_by_xpath("(//div[text()='"+cobertura_pais+"'])[2]").click()
        else:
            driver.find_element_by_xpath("//div[text()='"+cobertura_pais+"']").click()
        # state
        msg = "Coverage State"
        print(msg)
        driver.find_element_by_id("register-form_coverage_0_state").send_keys(cobertuda_estado)#//div[text()='"+cobertuda_estado+"']
        time.sleep(1)
        #la ciudad de contacto es igual al de cobertura?
        if local_estado == cobertuda_estado:
            if cobertuda_estado == local_ciudad:
                driver.find_element_by_xpath("(//div[text()='"+cobertuda_estado+"'])[3]").click()
            else:
                driver.find_element_by_xpath("(//div[text()='"+cobertuda_estado+"'])[2]").click()
        else:#(//div[@class='ant-select-item-option-content'])[2]|
            driver.find_element_by_xpath("//div[text()='" + cobertuda_estado + "']").click()

            #driver.find_element_by_xpath("//div[text()='" + cobertuda_estado + "']").click()
        msg = "inform. contacto"#####################################################
        print(msg)
        driver.find_element_by_id("register-form_contacts_0_first_name").send_keys(contact_nombre)
        driver.find_element_by_id("register-form_contacts_0_last_name").send_keys(contact_apellido)
        driver.find_element_by_id("register-form_contacts_0_email").send_keys(contact_correo)
        driver.find_element_by_id("register-form_contacts_0_mobile").send_keys(contact_mobile)
        msg = "main user"#####################################################
        print(msg)
        driver.find_element_by_id("register-form_first_name").send_keys(mainuser_nombre)
        driver.find_element_by_id("register-form_last_name").send_keys(mainuser_apellido)
        driver.find_element_by_id("register-form_mobile").send_keys(mainuser_mobile)
        driver.find_element_by_id("register-form_user_email").send_keys(mainuser_email)
        driver.find_element_by_id("register-form_password").send_keys(mainsuser_pass)
        driver.find_element_by_id("register-form_conf_password").send_keys(mainuser_confpass)

        driver.find_element_by_xpath("//button[@type='submit']").click()
        flag = 1
        print("Creando...")
        time.sleep(20)
        if(driver.find_element_by_xpath("//a[@href='/auth/register-3']").is_displayed()==True):
            msg = "FORMULARIO LLENADO NUMERO: " + str(r - 2)  #####################################################
            print(msg)
            xlc.writedata(path, 'tourop-create', r, 21, 'success')
            xlc.writedata(path, 'tourop-create', r, 22, '')
        pass
    except:
        value=driver.find_element_by_xpath("(//div[@role='alert'])[1]").get_attribute('innerHTML').strip()
        value=value[683:]
        value=value[:(len(value)-53)]
        print(value)
        xlc.writedata(path, 'tourop-create', r, 21, 'failed')
        if flag == 0:
            xlc.writedata(path, 'tourop-create', r, 22, msg)
        else:
            xlc.writedata(path, 'tourop-create', r, 22, value)
        print('FAILED!!')
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
        time.sleep(1)
        driver.find_element_by_id("back").click()
        time.sleep(3)
        #driver.find_element_by_xpath("(//button[contains(@class,'ant-btn ant-btn-default')])[1]").click()
        pass

print("FINISH!!")