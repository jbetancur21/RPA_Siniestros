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
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, id_HTML))
    )
    etiqueta.send_keys(valor)

service = Service(executable_path=r"C:/Users/Juan Pablo/Documents/RPA Formulario/chromedriver.exe")
chrome_options = Options()
chrome_options.add_argument("--window-size=1920,1080")
#chrome_options.add_argument("--headless")
driver = webdriver.Chrome(service=service, options=chrome_options)
try:
    driver.get("https://ecms.previsora.gov.co/appnet/UnityForm.aspx?d1=AfGMAHPe1S0R%2fcz2iAdBXSpJHIO0YgviS15UD0j27ERwLSQiFckq4W4e04Xjczr5hzxdnPy5kfHnZMpIVngsRo7o%2f%2bCFGiCSTAJQnX1YAoDne6%2bTEVUh8e4N%2fpDEokaHnA02hPiU9GS%2fzBLhUMjBb15PsKHMiTw7tqXHaFo508lP0tMeEtB2ZL8J%2bKJXi4vMzLSPBLwa1stB%2fpNzEXeBrNv9PcgenYvxGg4vEA1OL5RK3GrvEdkSFbqTRAiWfNqgQA%3d%3d")
    WebDriverWait(driver, 20).until(EC.alert_is_present())
    alert = driver.switch_to.alert
    alert.accept()
except Exception as e:
    print(f"Se presentó un error: {e}")
finally:
    driver.switch_to.frame('uf_hostframe')
    #Se Accede al input de Asegurado
    busqueda_elemento_id("asegurado_nombre",'asegurado11_input')
    #Se Accede al input de Tipo de documento
    busqueda_elemento_id("CC",'tipodocumento38_input')
    #Se Accede al input del documento
    busqueda_elemento_id("51651651",'idasegurado10_input')
    #Se Accede al input del numero de Poliza
    busqueda_elemento_id("54651651",'cuadrodetexto41_input')
    #Se Accede al input del Producto
    busqueda_elemento_id("AGRICOLA",'ramocomercial17_input')
    #Se Accede al input de la Sucursal
    busqueda_elemento_id("Medellin",'sucursal18_input')
    #Se Accede al DIV de los datos del contacto
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'sección42'))
        )
    etiqueta.click()
    time.sleep(2)
    #------------------------------------------------------------------------
    #Se Accede al input del nombre de contacto
    busqueda_elemento_id("contacto_nombre",'destinatarioexterno59_input')
    #Se Accede al input del telefono
    busqueda_elemento_id("3015141038",'cuadrodetexto60_input')
    #Se Accede al input del correo
    busqueda_elemento_id("contacto_email@gmail.com",'emailcontacto32_input')
    #Se Accede al input del correo por 2da vez
    busqueda_elemento_id("contacto_email@gmail.com",'cuadrodetexto40_input')
    #Se Accede al DIV de los datos del contacto
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'sección19'))
        )
    etiqueta.click()
    time.sleep(2)
    #------------------------------------------------------------------------
    #Se Accede al input de la fecha
    busqueda_elemento_id("09/02/2024",'fechasiniestro24_input')
    
    time.sleep(2)
    #Se Accede al input de la descripción
    busqueda_elemento_id("reclamo_descripcion",'cuadrodetextodelíneasmúltiples60_input')
    #------------------------------------------------------------------------
    #Se Accede al input del checkbox Autorización
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'casilladeverificación55_input'))
        )
    etiqueta.click()
    time.sleep(2)
    #------------------------------------------------------------------------
    #Se Accede al input del checkbox Entendido
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'casilladeverificación56_input'))
        )
    etiqueta.click()
    time.sleep(5)
    #------------------------------------------------------------------------
    #Se Accede al Botón de envíar
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/div/input'))
        )
    etiqueta.click()
    time.sleep(20)