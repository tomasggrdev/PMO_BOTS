from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import os
import json
import time
#variables

def descargar_notas():

    print("\ncore bot 1.0 notas :)\n")

    BASE_PATH = os.getcwd()
    DRIVER_PATH = BASE_PATH + "/conf/drivers/chromedriver.exe"
    CREDENTIALS_PATH = BASE_PATH + "/conf/credenciales.json"
    f = open(CREDENTIALS_PATH)
    CREDENTIALS = json.load(f)
    f.close()
    DOWNLOADS_PATH = BASE_PATH + "\\descargas\\notas"

    options = Options()
    #options.add_argument("--headless")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_experimental_option("prefs", {
      "download.default_directory": DOWNLOADS_PATH
      })

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.implicitly_wait(0.5)
    driver.maximize_window()

    driver.get("https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=679") #O1
    username=driver.find_element(By.ID,"username")
    password=driver.find_element(By.ID,"password")
    login = driver.find_element(By.ID,"loginbtn")
    username.send_keys(CREDENTIALS["USER"])
    password.send_keys(CREDENTIALS["PASS"])
    login.click()

    links = [
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=812",  # U1
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=823",  # U2
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=824",  # U3
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=825",  # U4
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=826",  # U5
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=827",  # U6
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=828",  # U7
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=829",  # U8
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=830",  # U9
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=813",  # U10
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=814",  # U11
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=815",  # U12
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=816",  # U13
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=817",  # U14
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=818",  # U15
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=819",  # U16
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=820",  # U17
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=821",  # U18
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=822",  # U19
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=831",  # U20
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=832",  # U21
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=833",  # U22
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=834",  # U23
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=835",  # U24
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=836",  # U25
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=837",  # U26
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=838",  # U27
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=839",  # U28
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=840",  # U29
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=841",  # U30
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=842",  # U31
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=843",  # U32
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=844",  # U33
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=845",  # U34
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=846",  # U35
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=847",  # U36
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=848",  # U37
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=849",  # Z1
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=858",  # Z2
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=859",  # Z3
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=860",  # Z4
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=861",  # Z5
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=862",  # Z6
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=863",  # Z7
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=864",  # Z8
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=865",  # Z9
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=850",  # Z10
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=851",  # Z11
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=852",  # Z12
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=853",  # Z13
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=854",  # Z14
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=855",  # Z15
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=856",  # Z16
        "https://lms.uis.edu.co/mintic2022/grade/export/xls/index.php?id=857"   # Z17
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