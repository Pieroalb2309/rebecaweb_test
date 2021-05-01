import time
import os
import excelfunctions as xlc
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

path = os.path.abspath("datatest.xlsx")
driver = xlc.start_webdriver()

driver.get("http://localhost:3000/auth/login")
driver.maximize_window()
rows=xlc.getrowcount(path,'poi-create')
print(driver.title)
#LOG IN---
msg="not logged"
driver.find_element_by_id("login-form_email").send_keys('sergio.sotelo@eulerinnovations.com')
driver.find_element_by_id("login-form_password").send_keys('eulerinnovations')
driver.find_element_by_xpath("//button[@type='submit']").click()
time.sleep(5)
print("logged!!")
driver.find_element_by_id("Places of interest").click()
#inicia el bucle
for r in range(3,rows+1):
    try:
        msg = "FORMULARIO NUMERO: " + str(r - 2)  #####################################################
        print(msg)
        flag=0
        nombre = xlc.readdata(path, 'poi-create', r, 1)
        mostrar = xlc.readdata(path, 'poi-create', r, 2)
        categoria = xlc.readdata(path, 'poi-create', r, 3)
        sinopsis = xlc.readdata(path, 'poi-create', r, 4)
        descripcion = xlc.readdata(path, 'poi-create', r, 5)
        pais = xlc.readdata(path, 'poi-create', r, 6)
        ciudad = xlc.readdata(path, 'poi-create', r, 7)
        direccion = xlc.readdata(path, 'poi-create', r, 8)
        lat = xlc.readdata(path, 'poi-create', r, 9)
        lon = xlc.readdata(path, 'poi-create', r, 10)  # string "0,1,2,3,4,5,6,7"
        c_tips = xlc.readdata(path, 'poi-create', r, 11)
        t_titulo = xlc.readdata(path, 'poi-create', r, 12)
        t_mensaje = xlc.readdata(path, 'poi-create', r, 13)
        t_tipo = xlc.readdata(path, 'poi-create', r, 14)
        t_notif = xlc.readdata(path, 'poi-create', r, 15)
        t_freq = xlc.readdata(path, 'poi-create', r, 16)
        t_url = xlc.readdata(path, 'poi-create', r, 17)
        c_recoms = xlc.readdata(path, 'poi-create', r, 18)
        r_titulo = xlc.readdata(path, 'poi-create', r, 19)
        r_mensaje = xlc.readdata(path, 'poi-create', r, 20)
        r_tipo = xlc.readdata(path, 'poi-create', r, 21)
        r_notif = xlc.readdata(path, 'poi-create', r, 22)
        r_freq = xlc.readdata(path, 'poi-create', r, 23)
        r_url = xlc.readdata(path, 'poi-create', r, 24)
        # ingresar al formulario
        driver.find_element_by_xpath("(//button[@type='button'])[2]").click()
        time.sleep(3)
        print("formulario")

        msg="basic info"
        print(msg)
        driver.find_element_by_xpath("//div[text()='Información Básica']").click()
        time.sleep(1)
        driver.find_element_by_id("register-form_name").send_keys(nombre)
        time.sleep(1)
        if mostrar == "Y":
            driver.find_element_by_id("register-form_show_in_map").click()

        driver.find_element_by_id("register-form_category").send_keys(categoria)
        time.sleep(1)
        driver.find_element_by_xpath("//div[text()='"+categoria+"']").click()#//div[text()='Restaurante']

        if sinopsis != "":
            driver.find_element_by_id("register-form_short_description-es").send_keys(sinopsis)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])[1]").click()
            #time.sleep(5)

        if descripcion != "":
            driver.find_element_by_id("register-form_description-es").send_keys(descripcion)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])[2]").click()
            #time.sleep(5)

        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
        time.sleep(1)
        ##
        cantidad=2
        tip_recom=1
        msg = "inform lugar"
        print(msg)
        driver.find_element_by_xpath("//div[text()='Información del lugar']").click()
        time.sleep(1)
        driver.find_element_by_id("register-form_country").send_keys(pais)
        time.sleep(1)
        driver.find_element_by_xpath("//div[text()='" + pais + "']").click()  # (//div[@class='ant-select-item-option-content'])[1]
        msg = "city"
        print(msg)
        time.sleep(1)
        driver.find_element_by_id("register-form_city").send_keys(ciudad)  # //div[text()='"+local_estado+"']
        time.sleep(1)
        value = (driver.find_element_by_xpath("(//div[@class='ant-select-item-option-content'])[1]").get_attribute(
            'innerHTML').strip())
        if value[0] == ciudad[0]:
            driver.find_element_by_xpath("(//div[@class='ant-select-item-option-content'])[1]").click()
        elif ciudad in value:
            driver.find_element_by_xpath("//div[text()='" + value + "']").click()
        else:
            driver.find_element_by_xpath("//div[text()='" + ciudad + "']").click()
        driver.find_element_by_id("register-form_address").send_keys(direccion)
        driver.find_element_by_id("register-form_latitude").send_keys(str(lat))
        driver.find_element_by_id("register-form_longitude").send_keys(str(lon))
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
        time.sleep(1)
        ##
        if c_tips == "Y":
            msg = "Tips"
            print(msg)
            driver.find_element_by_xpath("//div[text()='Tips de viaje']").click()
            time.sleep(1)
            driver.find_element_by_id("add-tip").click()
            time.sleep(1)
            cantidad=cantidad+1
            msg = "Tips_title"
            print(msg)
            driver.find_element_by_id("register-form_tips_0_title-es").send_keys(t_titulo)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])["+str(cantidad)+"]").click()
            #time.sleep(5)
            cantidad = cantidad + 1
            msg = "Tips_msg"
            print(msg)
            driver.find_element_by_id("register-form_tips_0_message-es").send_keys(t_mensaje)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])[" + str(cantidad) + "]").click()
            #time.sleep(5)
            msg = "Tips_type"
            print(msg)
            driver.find_element_by_id("register-form_tips_0_type").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='" + t_tipo + "']").click()
            if t_notif == "Y":
                msg = "Tips_notif"
                print(msg)
                driver.find_element_by_id("register-form_tips_0_is_push").click()
            msg = "Tips_freq"
            print(msg)
            driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
            time.sleep(1)
            driver.find_element_by_id("register-form_tips_0_frequency").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='" + t_freq + "']").click()
            msg = "Tips_url"
            print(msg)
            driver.find_element_by_id("register-form_tips_0_url").send_keys(t_url)

            tip_recom = tip_recom + 1
            driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
            time.sleep(1)
        ##
        if c_recoms == "Y":
            msg = "Recommendation"
            print(msg)
            driver.find_element_by_xpath("//div[text()='Información recomendada']").click()
            time.sleep(1)
            driver.find_element_by_id("add-recom").click()
            time.sleep(1)
            cantidad=cantidad+1
            msg = "Recommendation_title"
            print(msg)
            driver.find_element_by_id("register-form_recoms_0_title-es").click()
            driver.find_element_by_id("register-form_recoms_0_title-es").send_keys(r_titulo)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])[" + str(cantidad) + "]").click()
            #time.sleep(5)
            cantidad = cantidad + 1
            msg = "Recommendation_msg"
            print(msg)
            driver.find_element_by_id("register-form_recoms_0_message-es").click()
            driver.find_element_by_id("register-form_recoms_0_message-es").send_keys(r_mensaje)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])[" + str(cantidad) + "]").click()
            #time.sleep(5)
            msg = "Recommendation_type"
            print(msg)
            driver.find_element_by_id("register-form_recoms_0_type").click()
            time.sleep(1)
            driver.find_element_by_xpath("(// div[text() = '"+r_tipo+"'])["+str(tip_recom)+"]").click()
            if r_notif == "Y":
                msg = "Recommendation_notif"
                print(msg)
                driver.find_element_by_id("register-form_recoms_0_is_push").click()
            msg = "Recommendation_freq"
            print(msg)
            driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
            time.sleep(1)
            driver.find_element_by_id("register-form_recoms_0_frequency").click()
            time.sleep(1)
            driver.find_element_by_xpath("(//div[text()='" + r_freq + "'])["+str(tip_recom)+"]").click()
            msg = "Recommendation_url"
            print(msg)
            driver.find_element_by_id("register-form_recoms_0_url").send_keys(r_url)
            driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
            time.sleep(1)
        ##
        print("Subiendo...")
        driver.find_element_by_xpath("//button[text()='Enviar']").click()
        flag = 1
        time.sleep(10)
        if (driver.find_element_by_xpath("//h2[text()='Lugares de Interés']").click() == True):
            msg = "FORMULARIO LLENADO NUMERO: " + str(r - 2)  #####################################################
            print(msg)
            xlc.writedata(path, 'poi-create', r, 25, 'success')
            xlc.writedata(path, 'poi-create', r, 26, '')
        pass
    except:
        value = driver.find_element_by_xpath("(//div[@role='alert'])[1]").get_attribute('innerHTML').strip()
        value = value[683:]
        value = value[:(len(value) - 53)]
        print('failed')
        xlc.writedata(path, 'poi-create', r, 25,'failed')
        if flag == 0:
            xlc.writedata(path, 'poi-create', r, 26, msg)
        else:
            xlc.writedata(path, 'poi-create', r, 26, value)
        # 3 segundos para subir al inicio del formulario y click en el boton de regreso
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
        time.sleep(3)
        driver.find_element_by_xpath("(//button[@type='button'])[1]").click()
        pass
