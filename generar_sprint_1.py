from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

import os
import json

BASE_PATH = os.getcwd()
DRIVER_PATH = BASE_PATH + "/conf/drivers/chromedriver.exe"
CREDENTIALS_PATH = BASE_PATH + "/conf/credenciales.json"
f = open(CREDENTIALS_PATH)
CREDENTIALS = json.load(f)
f.close()

wb = Workbook()
dest_filename = 'sprint_1.xlsx'
ws1 = wb.active
ws1.title = "sprint_1"
ws1.append(["Codigo","Codigo_Id","Nombre_Tripulante", "Estado","Nota_Final","Grupo","Sprint"])

chrome_options = Options()
#chrome_options.add_argument("--disable-extensions")
#chrome_options.add_argument("--disable-gpu")
#chrome_options.add_argument("--headless")
#chrome_options.headless = True
driver = webdriver.Chrome(executable_path=DRIVER_PATH,options=chrome_options)
driver.implicitly_wait(0.5)


driver.get("https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37842&action=grading") #O1
username=driver.find_element(By.ID,"username")
password=driver.find_element(By.ID,"password")
login = driver.find_element(By.ID,"loginbtn")
username.send_keys(CREDENTIALS["USER"])
password.send_keys(CREDENTIALS["PASS"])
login.click()

links_sprint_1 = [
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37842&action=grading",#O1
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38028&action=grading",#O2
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37935&action=grading",#O3
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38214&action=grading",#O4
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38121&action=grading",#O5
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38307&action=grading",#O6
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38493&action=grading",#O7
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38679&action=grading",#O8
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38400&action=grading",#O9
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38586&action=grading",#10
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39330&action=grading",#O11
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39423&action=grading",#O12
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39144&action=grading",#O13
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39237&action=grading",#O14
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38865&action=grading",#O15
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38958&action=grading",#O16
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39051&action=grading",#O17
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38772&action=grading",#O18
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39516&action=grading",#O19
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39702&action=grading",#O20
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39795&action=grading",#O21
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39888&action=grading",#O22
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39981&action=grading",#O23
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39609&action=grading",#O24
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40260&action=grading",#O25
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40353&action=grading",#O26
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40074&action=grading",#O27
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40167&action=grading",#O28
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40446&action=grading",#O29
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40539&action=grading",#O30
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40818&action=grading",#O31
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40632&action=grading",#O32
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41097&action=grading",#O33
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40725&action=grading",#O34
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41190&action=grading",#O35
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40911&action=grading",#O36
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41004&action=grading",#O37
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41283&action=grading",#O38
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41469&action=grading",#O39
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41655&action=grading",#O40
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41562&action=grading",#O41
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41748&action=grading",#O42
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42027&action=grading",#O43
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42213&action=grading",#O44
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41934&action=grading",#O45
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41841&action=grading",#O46
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41376&action=grading",#O47
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42120&action=grading",#O48
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42399&action=grading",#O49
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42306&action=grading",#O50
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42492&action=grading",#O51
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42585&action=grading",#O52
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42678&action=grading",#O53
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42864&action=grading",#O54
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42771&action=grading",#O55
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42957&action=grading",#O56
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43050&action=grading",#O57
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43143&action=grading",#O58
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43236&action=grading",#O59

]

def read_sprint(links,numeroSprint,driver,hoja_excel):

    for i in range(59):
        contador = 1
        driver.get(links[i])
        flag = True
        a=4
        if i == 0:
            a = 5
        while flag == True:
            try:
                codigo = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[4]").text
                codigo_id = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[5]").text
                nombre = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[3]").text
                estado = driver.find_elements(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[6]/div")
                calificacion_final = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[15]").text
                #ws1 = wb.active
                
                for j in range(len(estado)):
                    if j==0:
                        hoja_excel.append([codigo,codigo_id,nombre,estado[j].text,calificacion_final,f"O{i+1}",f"s{numeroSprint}"])
                    else:
                        hoja_excel.append(["","","",estado[j].text,"","",""])


                contador+=1
            except:
                print(f"O{i+1} terminado")
                flag = False
        
  

read_sprint(links_sprint_1,1,driver,ws1)
wb.save(dest_filename)

driver.quit()