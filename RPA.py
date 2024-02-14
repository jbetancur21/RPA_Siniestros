from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import json
import time
import os
import shutil
from functions.functions import formato_fecha, descarga_archivos, mover_zip, crear_carpeta, cargar_archivos,comprimir
from datetime import datetime
import zipfile

def busqueda_elemento_id(valor,id_HTML):
    etiqueta = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, id_HTML))
    )
    etiqueta.send_keys(valor)

descarga_archivos()
mover_zip()

service = Service(executable_path=r"C:/Users/Juan Pablo/Documents/RPA Formulario/chromedriver.exe")
chrome_options = Options()
chrome_options.add_argument("--window-size=1920,1080")
#chrome_options.add_argument("--incognito")
#chrome_options.add_argument("--headless")

directorio = 'C:/Users/Juan Pablo/Documents/RPA Formulario/previsora_JSON/'
leidos = "leidos.txt"
fecha=datetime.now().strftime("%d_%m_%Y")
# Lista de nombres de archivos en el directorio
archivos_json = [archivo for archivo in os.listdir(directorio) if archivo.endswith('.json')]
crear_carpeta()

for nombre_archivo in archivos_json:
    ruta_archivo = os.path.join(directorio, nombre_archivo)
    #driver = webdriver.Chrome(service=service)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    with open(ruta_archivo, 'r', encoding='utf-8') as archivo_json:
        datos = json.load(archivo_json)
    for datos in datos:
        id= datos['id']
        asegurado_nombre= datos['asegurado_nombre']
        asegurado_tipodoc= datos['asegurado_tipodoc']
        asegurado_documento= datos['asegurado_documento']
        poliza_numero= datos['poliza_numero']
        poliza_producto= datos['poliza_producto']
        poliza_sucursal= datos['poliza_sucursal']
        contacto_nombre= datos['contacto_nombre']
        contacto_telefono= datos['contacto_telefono']
        contacto_email= datos['contacto_email']
        check_email= datos['contacto_email']
        reclamo_fecha= datos['reclamo_fecha']
        reclamo_descripcion= datos['reclamo_descripcion']
            
        with open(leidos, "r+") as archivo:
            contenido = archivo.read()
            if (str(id)+".json") not in contenido:
                try:
                    driver.get("https://ecms.previsora.gov.co/appnet/UnityForm.aspx?d1=AfGMAHPe1S0R%2fcz2iAdBXSpJHIO0YgviS15UD0j27ERwLSQiFckq4W4e04Xjczr5hzxdnPy5kfHnZMpIVngsRo7o%2f%2bCFGiCSTAJQnX1YAoDne6%2bTEVUh8e4N%2fpDEokaHnA02hPiU9GS%2fzBLhUMjBb15PsKHMiTw7tqXHaFo508lP0tMeEtB2ZL8J%2bKJXi4vMzLSPBLwa1stB%2fpNzEXeBrNv9PcgenYvxGg4vEA1OL5RK3GrvEdkSFbqTRAiWfNqgQA%3d%3d")
                    time.sleep(2)
                    WebDriverWait(driver, 15).until(EC.alert_is_present())
                    alert = driver.switch_to.alert
                    alert.accept()
                except Exception as e:
                    print(f"Se presentó un error: {e}")
                finally:
                    driver.switch_to.frame('uf_hostframe')

                    #Se Accede al input de Asegurado
                    busqueda_elemento_id(asegurado_nombre,'asegurado11_input')
                    #Se Accede al input de Tipo de documento
                    busqueda_elemento_id(asegurado_tipodoc,'tipodocumento38_input')
                    #Se Accede al input del documento
                    busqueda_elemento_id(asegurado_documento,'idasegurado10_input')
                    #Se Accede al input del numero de Poliza
                    busqueda_elemento_id(poliza_numero,'cuadrodetexto41_input')
                    #Se Accede al input del Producto
                    busqueda_elemento_id(poliza_producto,'ramocomercial17_input')
                    #Se Accede al input de la Sucursal
                    busqueda_elemento_id(poliza_sucursal,'sucursal18_input')
                    #Se Accede al DIV de los datos del contacto
                    etiqueta = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.ID, 'sección42'))
                        )
                    etiqueta.click()
                    time.sleep(2)
                    #------------------------------------------------------------------------
                    #Se Accede al input del nombre de contacto
                    busqueda_elemento_id(contacto_nombre,'destinatarioexterno59_input')
                    #Se Accede al input del telefono
                    busqueda_elemento_id(contacto_telefono,'cuadrodetexto60_input')
                    #Se Accede al input del correo
                    busqueda_elemento_id(contacto_email,'emailcontacto32_input')
                    #Se Accede al input del correo por 2da vez
                    busqueda_elemento_id(check_email,'cuadrodetexto40_input')
                    #Se Accede al DIV de los datos del contacto
                    etiqueta = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.ID, 'sección19'))
                        )
                    etiqueta.click()
                    time.sleep(2)
                    #------------------------------------------------------------------------
                    #Se Accede al input de la fecha
                    print(formato_fecha(reclamo_fecha))
                    script = f""" document.getElementById("fechasiniestro24_input").value = "{formato_fecha(reclamo_fecha)}";"""
                    driver.execute_script(script)

                    #busqueda_elemento_id(formato_fecha(reclamo_fecha),'fechasiniestro24_input')
                    time.sleep(2)
                    #Se Accede al input de la descripción
                    busqueda_elemento_id(reclamo_descripcion,'cuadrodetextodelíneasmúltiples60_input')
                    #------------------------------------------------------------------------
                    #Se Accede al input del checkbox Autorización
                    etiqueta = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.ID, 'casilladeverificación55_input'))
                        )
                    etiqueta.click()
                    time.sleep(2)
                    #------------------------------------------------------------------------
                    #Se Accede al input del checkbox Entendido
                    etiqueta = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.ID, 'casilladeverificación56_input'))
                        )
                    etiqueta.click()
                    time.sleep(5)
                    #------------------------------------------------------------------------
                    #Se Accede al Botón de envíar
                    etiqueta = WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/div/input'))
                        )
                    etiqueta.click()
                    time.sleep(5)
                    #---------------------PÁGINA DEL PANTALLAZO---------------------------------------------------
                    #Espera a que la página cargue y luego captura la pantalla
                    etiqueta = WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/section'))
                        )
                    
                    driver.save_screenshot(f"{fecha}/{id}.png")
                    with open(leidos, "a") as archivo:
                        archivo.write(str(id)+".json\n")
                    
                    driver.quit()
                    
comprimir(fecha)
cargar_archivos(fecha)
shutil.rmtree(directorio)