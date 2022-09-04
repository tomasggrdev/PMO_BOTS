from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import json
import os
import time


def descargar_grupos():

  BASE_PATH = os.getcwd()
  DRIVER_PATH = BASE_PATH + "/conf/drivers/chromedriver.exe"
  CREDENTIALS_PATH = BASE_PATH + "/conf/credenciales.json"
  f = open(CREDENTIALS_PATH)
  CREDENTIALS = json.load(f)
  f.close()
  DOWNLOADS_PATH = BASE_PATH + "\\descargas\\grupos"
  #DOWNLOADS_PATH = r"C:\Users\TOMASGONZALEZ\OneDrive - Universidad Industrial de Santander\MinTIC2_Ciclo3 - Archivos de MISION TIC - Coordinador de Proyectos\10in\11dbclt\1101grupos\Extraccion_grupos_dinamico"


  options = Options()
  #options.add_argument("--headless")
  options.add_experimental_option("excludeSwitches", ["enable-automation"])
  options.add_experimental_option('useAutomationExtension', False)
  options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOADS_PATH 
    })

  print("\ncore bot 1.0 notas :)\n")

  #set chpythonromodriver.exe path
  driver = webdriver.Chrome(executable_path=DRIVER_PATH,options = options)
  #implicit wait
  driver.implicitly_wait(0.5)
  #maximize browser
  driver.maximize_window()
  #https://lms.uis.edu.co/mintic2022/my/
  driver.get("https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=37817")
  username=driver.find_element(By.ID,"username")
  password=driver.find_element(By.ID,"password")
  login = driver.find_element(By.ID,"loginbtn")
  username.send_keys(CREDENTIALS["USER"])
  password.send_keys(CREDENTIALS["PASS"])
  login.click()

  links = [
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=37817", #O1
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38003", #O2
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=37910", #O3
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38189", #O4
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38096", #O5
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38282", #O6
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38468", #O7
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38654", #O8
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38375", #O9
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38561", #O10
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39305", #O11
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39398", #O12
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39119", #O13
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39212", #O14
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38840", #O15
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38933", #O16
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39026", #O17
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=38747", #O18
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39491", #O19
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39677", #O20
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39770", #O21
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39863", #O22
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39956", #O23
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=39584", #O24
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40235", #O25
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40328", #O26
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40049", #O27
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40142", #O28
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40421", #O29
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40514", #O30
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40793", #O31
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40607", #O32
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41072", #O33
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40700", #O34
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41165", #O35
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40886", #O36
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=40979", #O37
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41258", #O38
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41444", #O39
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41630", #O40
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41537", #O41
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41723", #O42
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42002", #O43
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42188", #O44
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41909", #O45
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41816", #O46
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=41351", #O47
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42095", #O48
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42374", #O49
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42281", #O50
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42467", #O51
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42560", #O52
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42653", #O53
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42839", #O54
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42746", #O55
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=42932", #O56
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=43025", #O57
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=43118", #O58
"https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=43211", #O59
]

  for i in range(len(links)):
      driver.get(links[i])
      descargarbtn = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div/ul/li[2]/div/form/button"); 
      descargarbtn.click()

  time.sleep(5)
  
if __name__ == '__main__':
    descargar_grupos()