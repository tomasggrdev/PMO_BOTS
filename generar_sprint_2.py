from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import json

BASE_PATH = os.getcwd()
DRIVER_PATH = BASE_PATH + "/conf/drivers/chromedriver.exe"
CREDENTIALS_PATH = BASE_PATH + "/conf/credenciales.json"
f = open(CREDENTIALS_PATH)
CREDENTIALS = json.load(f)
f.close()

wb = Workbook()
dest_filename = 'sprint_2.xlsx'
ws1 = wb.active
ws1.title = "sprint_2"
ws1.append(["Codigo","Codigo_Id","Nombre_Tripulante", "Estado","Nota_Final","Grupo","Sprint"])

chrome_options = Options()
#chrome_options.add_argument("--disable-extensions")
#chrome_options.add_argument("--disable-gpu")
#chrome_options.add_argument("--headless")
#chrome_options.headless = True
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.implicitly_wait(0.5)


driver.get("https://lms.uis.edu.co/mintic2022/my/") #O1
username=driver.find_element(By.ID,"username")
password=driver.find_element(By.ID,"password")
login = driver.find_element(By.ID,"loginbtn")
username.send_keys(CREDENTIALS["USER"])
password.send_keys(CREDENTIALS["PASS"])
login.click()

links_sprint_2 = [
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37870&action=grading",#O1
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38056&action=grading",#O2
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37963&action=grading",#O3
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38242&action=grading",#O4
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38149&action=grading",#O5
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38335&action=grading",#O6
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38521&action=grading",#O7
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38707&action=grading",#O8
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38428&action=grading",#O9
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38614&action=grading",#O10
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39358&action=grading",#O11
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39451&action=grading",#O12
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39172&action=grading",#O13
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39265&action=grading",#O14
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38893&action=grading",#O15
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38986&action=grading",#O16
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39079&action=grading",#O17
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38800&action=grading",#O18
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39544&action=grading",#O19
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39730&action=grading",#O20
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39823&action=grading",#O21
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39916&action=grading",#O22
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40009&action=grading",#O23
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39637&action=grading",#O24
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40288&action=grading",#O25
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40381&action=grading",#O26
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40102&action=grading",#O27
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40195&action=grading",#O28
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40474&action=grading",#O29
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40567&action=grading",#O30
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40846&action=grading",#O31
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40660&action=grading",#O32
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41125&action=grading",#O33
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40753&action=grading",#O34
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41218&action=grading",#O35
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40939&action=grading",#O36
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41032&action=grading",#O37
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41311&action=grading",#O38
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41497&action=grading",#O39
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41683&action=grading",#O40
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41590&action=grading",#O41
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41776&action=grading",#O42
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42055&action=grading",#O43
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42241&action=grading",#O44
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41962&action=grading",#O45
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41869&action=grading",#O46
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41404&action=grading",#O47
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42148&action=grading",#O48
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42427&action=grading",#O49
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42334&action=grading",#O50
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42520&action=grading",#O51
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42613&action=grading",#O52
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42706&action=grading",#O53
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42892&action=grading",#O54
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42799&action=grading",#O55
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42985&action=grading",#O56
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43078&action=grading",#O57
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43171&action=grading",#O58
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43264&action=grading",#O59

]

def read_sprint(links,numeroSprint,driver,hoja_excel):

    for i in range(59):
        contador = 1
        driver.get(links[i])
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
                        hoja_excel.append([codigo,codigo_id,nombre,estado[j].text,calificacion_final,f"O{i+1}",f"s{numeroSprint}"])
                    else:
                        hoja_excel.append(["","","",estado[j].text,"","",""])


                contador+=1
            except:
                print(f"O{i+1} terminado")
                flag = False
        
        

read_sprint(links_sprint_2,2,driver,ws1)
wb.save(dest_filename)

driver.quit()