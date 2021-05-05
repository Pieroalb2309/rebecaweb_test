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
rows = xlc.getrowcount(path, 'accomm-create')
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
driver.find_element_by_id("Services").click()
#inicia el bucle
for r in range(3,rows+1):
    try:
        msg = "FORMULARIO NUMERO: " + str(r - 2)  #####################################################
        print(msg)
        flag = 0
        nombre = xlc.readdata(path, 'accomm-create', r, 1)
        proveedor = xlc.readdata(path, 'accomm-create', r, 2)
        nummax = xlc.readdata(path, 'accomm-create', r, 3)
        moneda = xlc.readdata(path, 'accomm-create', r, 4)
        costo = xlc.readdata(path, 'accomm-create', r, 5)
        sinopsis = xlc.readdata(path, 'accomm-create', r, 6)
        descripcion = xlc.readdata(path, 'accomm-create', r, 7)
        f_inic = xlc.readdata(path, 'accomm-create', r, 8)
        f_final = xlc.readdata(path, 'accomm-create', r, 9)
        checkbox = xlc.readdata(path, 'accomm-create', r, 10)#string "0,1,2,3,4,5,6,7"
        h_inic = xlc.readdata(path, 'accomm-create', r, 11)
        h_final = xlc.readdata(path, 'accomm-create', r, 12)
        pais = xlc.readdata(path, 'accomm-create', r, 13)
        ciudad = xlc.readdata(path, 'accomm-create', r, 14)
        direccion = xlc.readdata(path, 'accomm-create', r, 15)
        lat = xlc.readdata(path, 'accomm-create', r, 16)
        lon = xlc.readdata(path, 'accomm-create', r, 17)
        tiprecom = xlc.readdata(path, 'accomm-create', r, 18)
        titulo = xlc.readdata(path, 'accomm-create', r, 19)
        mensaje = xlc.readdata(path, 'accomm-create', r, 20)
        tipo = xlc.readdata(path, 'accomm-create', r, 21)
        notificacion = xlc.readdata(path, 'accomm-create', r, 22)
        frecuencia = xlc.readdata(path, 'accomm-create', r, 23)
        url = xlc.readdata(path, 'accomm-create', r, 24)
        c_pregunta = xlc.readdata(path, 'accomm-create', r, 25)
        pregunta = xlc.readdata(path, 'accomm-create', r, 26)
        respuesta = xlc.readdata(path, 'accomm-create', r, 27)
        video_ref = xlc.readdata(path, 'accomm-create', r, 28)
        c_distribucion = xlc.readdata(path, 'accomm-create', r, 29)
        des_piso = xlc.readdata(path, 'accomm-create', r, 30)
        tipo_espacio = xlc.readdata(path, 'accomm-create', r, 31)
        des_espacio = xlc.readdata(path, 'accomm-create', r, 32)
        des_item = xlc.readdata(path, 'accomm-create', r, 33)
        tipo_item = xlc.readdata(path, 'accomm-create', r, 34)
        cant_item = xlc.readdata(path, 'accomm-create', r, 35)
        precio_item = xlc.readdata(path, 'accomm-create', r, 36)
        # ingresar al formulario
        time.sleep(3)
        driver.find_element_by_xpath("(//button[@type='button'])[3]").click()
        time.sleep(1)
        driver.find_element_by_id("Accommodation").click()
        time.sleep(3)
        print("formulario")
        msg = "Inform. Basica"
        print(msg)
        driver.find_element_by_xpath("//div[text()='Información Básica']").click()####################
        time.sleep(1)
        driver.find_element_by_xpath("//input[@class='ant-input mr-lg-4']").send_keys(nombre)
        time.sleep(1)
        driver.find_element_by_id("register-form_provider").send_keys(proveedor)
        time.sleep(1)
        driver.find_element_by_xpath("//div[text()='"+proveedor+"']").click()
        driver.find_element_by_id("register-form_max_persons").send_keys(nummax)
        driver.find_element_by_id("register-form_currency").click()
        time.sleep(1)
        driver.find_element_by_xpath("//div[@title='"+moneda+"']").click()
        driver.find_element_by_id("register-form_cost").send_keys(str(costo))
        driver.find_element_by_id("register-form_activity_short_description-es").send_keys(sinopsis)
        #driver.find_element_by_xpath("(//button[text()='Traducción'])[1]").click()
        time.sleep(1)
        driver.find_element_by_id("register-form_activity_description-es").send_keys(descripcion)
        #driver.find_element_by_xpath("(//button[text()='Traducción'])[2]").click()
        time.sleep(1)
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
        time.sleep(1)
        ##
        msg = "Fecha de Planificacion"
        print(msg)
        driver.find_element_by_xpath("//div[text()='Fechas de planificacion']").click()#################
        time.sleep(1)
        driver.find_element_by_id("add-dates").click()
        time.sleep(1)
        driver.find_element_by_xpath("//input[@placeholder='Fecha inicial']").send_keys(f_inic)
        driver.find_element_by_xpath("//input[@placeholder='Fecha final']").send_keys(f_final)
        check = xlc.weeklist(checkbox)
        if check:
            for i in range(len(check)):
                driver.find_element_by_xpath("//span[text()='"+check[i]+"']").click()
        driver.find_element_by_id("add-hours").click()
        time.sleep(1)
        driver.find_element_by_id("register-form_dates_0_hour_list_0_timePick").click()
        time.sleep(1)
        driver.find_element_by_id("register-form_dates_0_hour_list_0_timePick").send_keys(str(h_inic))
        time.sleep(1)
        driver.find_element_by_id("register-form_dates_0_hour_list_0_timePick").send_keys("\n")
        time.sleep(1)
        driver.find_element_by_xpath("(//div[@class='ant-picker-input ant-picker-input-active']//input)[2]").send_keys(str(h_final)+"\n")
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
        time.sleep(1)
        ##
        msg = "Inform Lugar"
        print(msg)
        driver.find_element_by_xpath("//div[text()='Información del lugar']").click()
        time.sleep(1)
        driver.find_element_by_id("register-form_country").send_keys(pais)
        time.sleep(1)
        driver.find_element_by_xpath("//div[text()='" + pais + "']").click()  # (//div[@class='ant-select-item-option-content'])[1]
        msg = "city"
        print(msg)
        driver.find_element_by_id("register-form_city").send_keys(ciudad)  # //div[text()='"+local_estado+"']
        time.sleep(1)
        value = (driver.find_element_by_xpath("(//div[@class='ant-select-item-option-content'])[1]").get_attribute(
            'innerHTML').strip())
        if value[0] == ciudad[0]:
            driver.find_element_by_id("register-form_city").send_keys("\n")
        elif ciudad in value:
            driver.find_element_by_xpath("//div[text()='" + value + "']").click()
        else:
            driver.find_element_by_xpath("//div[text()='" + ciudad + "']").click()
        driver.find_element_by_id("register-form_address").send_keys(direccion)
        driver.find_element_by_xpath("(//input[@class='ant-input'])[4]").send_keys(str(lat))
        driver.find_element_by_xpath("(//input[@class='ant-input'])[5]").send_keys(str(lon))
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
        time.sleep(1)
        ##
        cantidad=2
        if tiprecom == "tips":
            msg = "Tips"
            print(msg)
            cantidad=cantidad+1
            driver.find_element_by_xpath("//div[text()='Tips de viaje']").click()
            time.sleep(1)
            driver.find_element_by_id("add-tip").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_tips_0_title-es").send_keys(titulo)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])["+str(cantidad)+"]").click()
            time.sleep(1)
            cantidad=cantidad+1
            driver.find_element_by_id("register-form_tips_0_message-es").send_keys(mensaje)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])["+str(cantidad)+"]").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_tips_0_type").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='"+tipo+"']").click()
            time.sleep(1)
            if notificacion == "Y":
                driver.find_element_by_id("register-form_tips_0_is_push").click()
            driver.find_element_by_id("register-form_tips_0_frequency").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='"+frecuencia+"']").click()
            driver.find_element_by_id("register-form_tips_0_url").send_keys(url)
        elif tiprecom == "recoms":
            msg = "Reccomendation"
            print(msg)
            cantidad=cantidad+1
            driver.find_element_by_xpath("//div[text()='Información recomendada']").click()
            time.sleep(1)
            driver.find_element_by_id("add-recom").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_recoms_0_title-es").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_recoms_0_title-es").sendKeys(titulo)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])["+str(cantidad)+"]").click()
            time.sleep(1)
            cantidad=cantidad+1
            driver.find_element_by_id("register-form_recoms_0_message-es").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_recoms_0_message-es").send_keys(mensaje)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])["+str(cantidad)+"]").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_recoms_0_type").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='"+tipo+"']").click()
            time.sleep(1)
            if notificacion == "Y":
                driver.find_element_by_id("register-form_recoms_0_is_push").click()
            driver.find_element_by_id("register-form_recoms_0_frequency").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='"+frecuencia+"']").click()
            driver.find_element_by_id("register-form_recoms_0_url").send_keys(url)
        ##
        if c_pregunta == "Y":
            msg = "Questions"
            print(msg)
            driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
            time.sleep(1)
            cantidad = cantidad + 1
            driver.find_element_by_xpath("//div[text()='FAQ']").click()
            time.sleep(1)
            driver.find_element_by_id("add-quest").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_questions_0_quest").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_questions_0_quest").send_keys(pregunta+"\n")
            time.sleep(1)
            #driver.find_element_by_xpath("//div[text()='"+pregunta+"']").click()
            #time.sleep(1)
            driver.find_element_by_id("register-form_questions_0_answer-es").click()
            driver.find_element_by_id("register-form_questions_0_answer-es").send_keys(respuesta)
            #driver.find_element_by_xpath("(//button[text()='Traducción'])[" + str(cantidad) + "]").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_questions_0_video_ref").click()
            driver.find_element_by_id("register-form_questions_0_video_ref").send_keys(video_ref)
            #driver.find_element_by_xpath("(//div[@class='ant-select-item-option-content'])[4]").click()

        ##
        if c_distribucion == "Y":
            msg = "Distribution"
            print(msg)
            driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
            time.sleep(2)
            driver.find_element_by_xpath("//div[text()='Distribución del alojamiento']").click()
            time.sleep(1)
            cantidad=cantidad+1
            driver.find_element_by_id("add-distr").click()
            time.sleep(1)
            if des_piso != "":
                driver.find_element_by_id("register-form_floors_0_floor_description-es").send_keys(des_piso)

            driver.find_element_by_xpath("//button[text()='Agregar espacio']").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_floors_0_spaces_0_space").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='"+tipo_espacio+"']").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_floors_0_spaces_0_space_description-es").send_keys(des_espacio)
            time.sleep(1)
            driver.find_element_by_xpath("//button[text()='Agregar objeto']").click()
            time.sleep(1)
            if des_espacio != "":
                driver.find_element_by_id("register-form_floors_0_spaces_0_items_0_description-es").send_keys(des_item)
                time.sleep(1)

            driver.find_element_by_id("register-form_floors_0_spaces_0_items_0_entity").click()
            time.sleep(1)
            driver.find_element_by_xpath("//div[text()='"+tipo_item+"']").click()
            time.sleep(1)
            driver.find_element_by_id("register-form_floors_0_spaces_0_items_0_quantity").send_keys(cant_item)
            time.sleep(1)
            driver.find_element_by_id("register-form_floors_0_spaces_0_items_0_value").send_keys(str(precio_item))
            time.sleep(1)
        ##
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.END)
        time.sleep(3)
        print("Subiendo....")
        driver.find_element_by_xpath("//button[text()='Enviar']").click()
        flag=1
        time.sleep(30)
        if (driver.find_element_by_xpath("//h2[text()='Servicios']").click() == True):
            msg = "FORMULARIO LLENADO NUMERO: " + str(r - 2)  #####################################################
            print(msg)
            xlc.writedata(path, "accomm-create", r, 36, 'success')
            xlc.writedata(path, "accomm-create", r, 37, '')
            xlc.writedata(path, "accomm-create", r, 38, d1)
        #driver.find_element_by_xpath("(//button[@type='button'])[1]").click()

        pass
    except:
        #value = driver.find_element_by_xpath("(//div[@role='alert'])[1]").get_attribute('innerHTML').strip()
        #value = value[683:]
        #value = value[:(len(value) - 53)]
        # xlc.writedata(path, 'login', r, 3, msg)
        print('failed')
        xlc.writedata(path, "accomm-create", r, 39, d1)
        xlc.writedata(path, "accomm-create", r, 37, 'failed')
        if flag == 0:
            xlc.writedata(path, "accomm-create", r, 38, msg)
        else:
            xlc.writedata(path, "accomm-create", r, 38, 'Selenium probl.')
        # 3 segundos para subir al inicio del formulario y click en el boton de regreso
        driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
        time.sleep(3)
        driver.find_element_by_xpath("(//button[@type='button'])[1]").click()
        pass

print("end")
# driver.close()