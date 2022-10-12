from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import json
import os
import time


def descargar_grupos():
  print("\ncore bot 1.0 grupos :)\n")
  BASE_PATH = os.getcwd()
  CREDENTIALS_PATH = BASE_PATH + "/conf/credenciales.json"
  DOWNLOADS_PATH = BASE_PATH + "\\descargas\\grupos"
  f = open(CREDENTIALS_PATH)
  CREDENTIALS = json.load(f)
  f.close()

  options = Options()
  # chrome_options.add_argument("--disable-extensions")
  # chrome_options.add_argument("--disable-gpu")
  # chrome_options.add_argument("--headless")
  # chrome_options.headless = True
  options.add_experimental_option("excludeSwitches", ["enable-automation"])
  options.add_experimental_option('useAutomationExtension', False)
  options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOADS_PATH
  })
  driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options = options)
  driver.implicitly_wait(0.5)
#driver.maximize_window()

  driver.get("https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=37817")
  username=driver.find_element(By.ID,"username")
  password=driver.find_element(By.ID,"password")
  login = driver.find_element(By.ID,"loginbtn")
  username.send_keys(CREDENTIALS["USER"])
  password.send_keys(CREDENTIALS["PASS"])
  login.click()



  links = [
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52577",  # U1
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52506",  # U2
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52364",  # U3
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52435",  # U4
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52293",  # U5
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52222",  # U6
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52151",  # U7
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52009",  # U8
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=52080",  # U9
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51654",  # U10
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51725",  # U11
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51796",  # U12
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51867",  # U13
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51938",  # U14
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51583",  # U15
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51512",  # U16
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51441",  # U17
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51370",  # U18
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51299",  # U19
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51228",  # U20
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51157",  # U21
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51086",  # U22
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=51015",  # U23
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50944",  # U24
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50660",  # U25
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50873",  # U26
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50802",  # U27
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50589",  # U28
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50731",  # U29
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50518",  # U30
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50447",  # U31
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50376",  # U32
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50305",  # U33
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50234",  # U34
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50021",  # U35
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50092",  # U36
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=50163",  # U37
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=48338",  # Z1
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=48536",  # Z2
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=48635",  # Z3
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=48734",  # Z4
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=48437",  # Z5
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=48833",  # Z6
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=48932",  # Z7
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49031",  # Z8
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49229",  # Z9
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49130",  # Z10
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49823",  # Z11
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49724",  # Z12
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49625",  # Z13
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49922",  # Z14
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49526",  # Z15
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49328",  # Z16
    "https://lms.uis.edu.co/mintic2022/mod/choice/report.php?id=49427"   # Z17
  ]

  for i in range(len(links)):
      driver.get(links[i])
      descargarbtn = driver.find_element(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div/ul/li[2]/div/form/button"); 
      descargarbtn.click()

  time.sleep(5)
  
if __name__ == '__main__':
    descargar_grupos()