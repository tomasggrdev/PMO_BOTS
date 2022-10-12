from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import os
import json

BASE_PATH = os.getcwd()
CREDENTIALS_PATH = BASE_PATH + "/conf/credenciales.json"
f = open(CREDENTIALS_PATH)
CREDENTIALS = json.load(f)
f.close()

chrome_options = Options()
#chrome_options.add_argument("--disable-extensions")
#chrome_options.add_argument("--disable-gpu")
#chrome_options.add_argument("--headless")
#chrome_options.headless = True
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.implicitly_wait(0.5)

wb = Workbook()
dest_filename = 'sprint_1.xlsx'
ws1 = wb.active
ws1.title = "sprint_1"
ws1.append(["Codigo","Codigo_Id","Nombre_Tripulante", "Estado","Nota_Final","Grupo","Sprint"])




driver.get("https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48377&action=grading") #Z1
username=driver.find_element(By.ID,"username")
password=driver.find_element(By.ID,"password")
login = driver.find_element(By.ID,"loginbtn")
username.send_keys(CREDENTIALS["USER"])
password.send_keys(CREDENTIALS["PASS"])
login.click()



links_sprint_1 = [
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52595&action=grading",'U1'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52524&action=grading",'U2'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52382&action=grading",'U3'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52453&action=grading",'U4'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52311&action=grading",'U5'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52240&action=grading",'U6'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52169&action=grading",'U7'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52683&action=grading",'U8'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52098&action=grading",'U9'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51672&action=grading",'U10'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51743&action=grading",'U11'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51814&action=grading",'U12'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51885&action=grading",'U13'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51956&action=grading",'U14'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51601&action=grading",'U15'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51530&action=grading",'U16'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51459&action=grading",'U17'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51388&action=grading",'U18'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52684&action=grading",'U19'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51246&action=grading",'U20'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51175&action=grading",'U21'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51104&action=grading",'U22'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51033&action=grading",'U23'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50962&action=grading",'U24'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50678&action=grading",'U25'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50891&action=grading",'U26'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50820&action=grading",'U27'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50607&action=grading",'U28'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50749&action=grading",'U29'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50536&action=grading",'U30'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50465&action=grading",'U31'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50394&action=grading",'U32'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50323&action=grading",'U33'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50252&action=grading",'U34'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50039&action=grading",'U35'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50110&action=grading",'U36'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52685&action=grading",'U37'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48377&action=grading",'Z1'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48575&action=grading",'Z2'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48674&action=grading",'Z3'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48773&action=grading",'Z4'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48476&action=grading",'Z5'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48872&action=grading",'Z6'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48971&action=grading",'Z7'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49070&action=grading",'Z8'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49268&action=grading",'Z9'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49169&action=grading",'Z10'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49862&action=grading",'Z11'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49763&action=grading",'Z12'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52686&action=grading",'Z13'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49961&action=grading",'Z14'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49565&action=grading",'Z15'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49367&action=grading",'Z16'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49466&action=grading",'Z17']
]

def read_sprint(links,numeroSprint,driver,hoja_excel):

    for i in range(len(links)):
        contador = 1
        driver.get(links[i][0])
        flag = True

        while flag == True:
            try:

                codigo = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[4]").text
                codigo_id = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[5]").text
                nombre = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[3]").text
                estado = driver.find_elements(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[6]/div")
                calificacion_final = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[15]").text
                #ws1 = wb.active
                
                for j in range(len(estado)):
                    if j==0:
                        hoja_excel.append([codigo,codigo_id,nombre,estado[j].text,calificacion_final,links[i][1],f"s{numeroSprint}"])
                    else:
                        hoja_excel.append(["","","",estado[j].text,"","",""])


                contador+=1
            except:
                print(f"{links[i][1]} terminado")
                flag = False
        
  

read_sprint(links_sprint_1,1,driver,ws1)
wb.save(dest_filename)

driver.quit()