from ast import If
import configparser
from re import T
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd
from datetime import date
import pyodbc
from fast_to_sql import fast_to_sql as fts

#---------------------------------------------------------------------------
# Conecta a SQL
def SQL_conexion (server, database):
    SQLConn = pyodbc.connect("Driver={ODBC Driver 17 for SQL Server} ;"
                     "Server=" + server + ";"
                     "Database=" + database + ";"
                     "Trusted_Connection=yes;")

    return SQLConn  #.cursor()
#-----------------------------------------------------
def get_renappo (URL, driver):

    driver.get(URL)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME,"paging_numbers")))
    Select (driver.find_element(by=By.NAME, value='table-padron_length')).select_by_index(3)
    page_counts = int(driver.find_element(by=By.XPATH, value='//*[@id="table-padron_paginate"]/ul/li[7]/a').text)
    print(f"Found - {page_counts} - pages")

    # Crea el dataframe para guardar los datos
    Renappo = pd.DataFrame (columns=['Nro','RazonSocial','Provincia','Localidad','CUIT','Actividades','Tipo'])
    pagina = 0

    # Recorre página por página y agrega los datos al DF
    for x in range(page_counts):
        pagina += 1
        print("Pagina n° ",pagina)
        element = driver.find_elements(by=By.XPATH, value='//*[@class="table table-hover display dataTable no-footer"]/tbody/tr')

        for i in element:
                
            fila = i.find_elements(by=By.TAG_NAME, value="td")
            Renappo = pd.concat ([Renappo,pd.Series([fila[0].text,fila[1].text,fila[2].text,fila[3].text ,fila[4].text.replace("-",""),fila[5].text,fila[6].text],index = Renappo.columns).to_frame().T])	
                
            print(fila[0].text,fila[1].text,fila[2].text,fila[3].text ,fila[4].text.replace("-",""),fila[5].text,fila[6].text)
        
        # Ir a la siguiente pagina
        if pagina < page_counts:
            LI = driver.find_elements(by=By.CLASS_NAME,value="paginate_button")
            for linea in LI:
                if linea.text != "…" :
                    if int(linea.text) > pagina:
                        linea.find_element(by=By.TAG_NAME, value="a").click()
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME,"pagination")))
                        break
        else:
            break

    return Renappo
#---------------------------------------------------------------------------------
def graba_sql (df, Conn):

    cursor = Conn.cursor()

    cursor.execute("SET ANSI_WARNINGS  OFF")
    cursor.commit()


    create_statement = fts.fast_to_sql(df, 'dbo.Renappo', Conn, if_exists='replace')
    Conn.commit()


#-----------------------------------------------------
cp = configparser.ConfigParser()
cp.read("config.ini")
URL = cp["DEFAULT"]["URL"]
Server_Origen = cp["DEFAULT"]["server_origen"]
Base_Origen= cp["DEFAULT"]["base_origen"]

Conn = SQL_conexion(Server_Origen, Base_Origen)

# Configura Chromedriver
options = webdriver.ChromeOptions()
options.headless = False
preferences = { "download.directory_upgrade": True,
                "safebrowsing_for_trusted_sources_enabled": False,
                "safebrowsing.enabled": False,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "download.prompt_for_download": False }
options.add_argument("--ignore-certificate-errors")
options.add_experimental_option("prefs", preferences)
options.add_experimental_option("excludeSwitches", ['enable-automation'])
options.add_argument('--kiosk-printing')
options.add_argument('--disable-gpu')
options.add_argument('--disable-software-rasterizer')

driver = webdriver.Chrome(options=options)

df= get_renappo(URL, driver)

graba_sql(df, Conn)
