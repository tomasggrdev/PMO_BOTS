from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
import json
import time
#variables

def descargar_notas():
    BASE_PATH = os.getcwd()
    DRIVER_PATH = BASE_PATH + "/conf/drivers/chromedriver.exe"
    CREDENTIALS_PATH = BASE_PATH + "/conf/credenciales.json"
    f = open(CREDENTIALS_PATH)
    CREDENTIALS = json.load(f)
    f.close()
    DOWNLOADS_PATH = BASE_PATH + "\\descargas\\notas"
    #DOWNLOADS_PATH = r"C:\Users\TOMASGONZALEZ\OneDrive - Universidad Industrial de Santander\MinTIC2_Ciclo3 - Archivos de MISION TIC - Coordinador de Proyectos\10in\11dbclt\1100notas\notas_dinamico"


    options = Options()
    #options.add_argument("--headless")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_experimental_option("prefs", {
      "download.default_directory": DOWNLOADS_PATH 
      })

    #set chromodriver.exe path
    print("\ncore bot 1.0 grupos :)\n")
    driver = webdriver.Chrome(executable_path=DRIVER_PATH, options = options)
    #implicit wait
    driver.implicitly_wait(0.5)
    #maximize browser
    driver.maximize_window()
    #chrome_Options().add_argument("--headless")

    #launch URL
    driver.get("https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=679") #O1
    username=driver.find_element(By.ID,"username")
    password=driver.find_element(By.ID,"password")
    login = driver.find_element(By.ID,"loginbtn")
    username.send_keys(CREDENTIALS["USER"])
    password.send_keys(CREDENTIALS["PASS"])
    login.click()


    links = [
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=679", #O1
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=690", #O2
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=692", #O3
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=693", #O4
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=694", #O5
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=695", #O6
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=696", #O7
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=697", #O8
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=698", #O9
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=680", #O10
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=681", #O11
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=682", #O12
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=683", #O13
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=684", #O14
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=685", #O15
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=686", #O16
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=687", #O17
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=688", #O18
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=689", #O19
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=691", #O20
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=529", #O21
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=530", #O22
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=531", #O23
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=532", #O24
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=533", #O25
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=534", #O26
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=535", #O27
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=536", #O28
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=537", #O29
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=518", #O30
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=519", #O31
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=520", #O32
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=521", #O33
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=522", #O34
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=523", #O35
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=524", #O36
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=525", #O37
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=526", #O38
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=527", #O39
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=508", #O40
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=509", #O41
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=510", #O42
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=511", #O43
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=512", #O44
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=513", #O45
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=514", #O46
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=515", #O47
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=516", #O48
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=517", #O49
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=498", #O50
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=499", #O51
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=500", #O52
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=501", #O53
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=502", #O54
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=503", #O55
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=504", #O56
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=505", #O57
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=506", #O58
"https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=507", #O59
]

    for i in range(len(links)):
        driver.get(links[i])
        descargarbtn = driver.find_element(By.ID,"id_submitbutton")
        descargarbtn.click()

    time.sleep(5)

    #close browser
    #driver.quit()

if __name__ == '__main__':
    descargar_notas()