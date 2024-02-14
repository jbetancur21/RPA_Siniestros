from dateutil import parser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException, NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import shutil
import zipfile
import os
import pyautogui
import json
from datetime import datetime


service = Service(executable_path=r"C:/Users/Juan Pablo/Documents/RPA Formulario/chromedriver.exe")

def formato_fecha(fecha):
    try:
        parsed_fecha = parser.parse(fecha)
        return parsed_fecha.strftime('%d/%m/%Y')
    except ValueError:
        return "Formato no detectado"
    

def descarga_archivos():
    data = credenciales()
    chrome_options = Options()
    chrome_options.add_argument("--window-size=1800,800")
    #chrome_options.add_argument("--disable-gpu")
    #chrome_options.add_argument("--headless")
     
    driver = webdriver.Chrome(service=service, options=chrome_options)
    #driver = webdriver.Chrome(service=service)
    driver.get(data['directorio_descarga'])
    
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]"))
    )
    etiqueta.send_keys(data['correo'])
    
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input"))
            )
    etiqueta.click()
    
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div/div[2]/input"))
    )
    etiqueta.send_keys(data['contraseña'])
    
    time.sleep(5)
    
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[5]/div/div/div/div"))
            )
    etiqueta.click()
        
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div[1]"))
            )
    etiqueta.click()  
    
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div[2]/div[1]/div/div/div/div/div/div[1]/div[5]"))
                
                
            )
    etiqueta.click()
    time.sleep(15)
    driver.quit()
    
        
def mover_zip():
    data = credenciales()
    previsora_JSON = f"{data['descarga_pc']}previsora_JSON.zip"
    RPA_JSON = 'C:/Users/Juan Pablo/Documents/RPA Formulario/'

    shutil.move(previsora_JSON, data['ruta_raiz'])
    
    with zipfile.ZipFile(data['ruta_raiz']+"previsora_JSON.zip", 'r') as archivo_zip:
        archivo_zip.extractall(data['ruta_raiz'])
        
    os.remove(data['ruta_raiz']+"previsora_JSON.zip")
    
def cargar_archivos(fecha):
    data = credenciales()
    chrome_options = Options()
    #chrome_options.add_argument("--headless")
    #chrome_options.add_argument("--disable-gpu") 
    chrome_options.add_argument("--window-size=1800,800")
    
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.get(data['directorio_carga'])
    
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]"))
    )
    etiqueta.send_keys(data['correo'])
    
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input"))
            )
    etiqueta.click()
    
    etiqueta = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div/div[2]/input"))
    )
    etiqueta.send_keys(data['contraseña'])
    
    time.sleep(5)
    
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[5]/div/div/div/div"))
            )
    etiqueta.click()
        
    etiqueta = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div[1]"))
            )
    etiqueta.click()
    
    etiqueta = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div[1]/div/nav/div[1]/div/div/span"))
            )
    etiqueta.click()
    time.sleep(5)
    
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div/div/div/div/ul/li[3]"))
            )
    etiqueta.click()
    time.sleep(2)
    
    etiqueta = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/input"))
            )
    etiqueta.send_keys(f"{data['ruta_raiz']}{fecha}.zip")
    driver.implicitly_wait(5)
    time.sleep(5)
    pyautogui.press('esc')
    driver.quit()
    
    time.sleep(10)
    
def crear_carpeta():
    if not os.path.exists(str(datetime.now().strftime("%d_%m_%Y"))):
        os.makedirs(str(datetime.now().strftime("%d_%m_%Y")))
        
def comprimir(comprimido):
    zip = zipfile.ZipFile(f"{comprimido}.zip", 'w')
    for archivo in os.listdir(comprimido):
        zip.write(os.path.join(comprimido, archivo))
        
def credenciales():
    data = "C:/Users/Juan Pablo/Documents/RPA Formulario/credenciales.json"
    with open(data, 'r', encoding='utf-8') as archivo_json:
        variables = json.load(archivo_json)
    return(variables)