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
# chrome_options.add_argument("--disable-extensions")
# chrome_options.add_argument("--disable-gpu")
# chrome_options.add_argument("--headless")
# chrome_options.headless = True
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.implicitly_wait(0.5)

wb = Workbook()
dest_filename = 'sprint_2.xlsx'
ws1 = wb.active
ws1.title = "sprint_2"
ws1.append(["Codigo", "Codigo_Id", "Nombre_Tripulante", "Estado", "Nota_Final", "Grupo", "Sprint"])

driver.get("https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48377&action=grading")  # Z1
username = driver.find_element(By.ID, "username")
password = driver.find_element(By.ID, "password")
login = driver.find_element(By.ID, "loginbtn")
username.send_keys(CREDENTIALS["USER"])
password.send_keys(CREDENTIALS["PASS"])
login.click()

links_sprint_1 = [
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52610&action=grading", 'U1'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52539&action=grading", 'U2'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52397&action=grading", 'U3'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52468&action=grading", 'U4'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52326&action=grading", 'U5'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52255&action=grading", 'U6'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52184&action=grading", 'U7'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52042&action=grading", 'U8'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=52113&action=grading", 'U9'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51687&action=grading", 'U10'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51758&action=grading", 'U11'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51829&action=grading", 'U12'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51900&action=grading", 'U13'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51971&action=grading", 'U14'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51616&action=grading", 'U15'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51545&action=grading", 'U16'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51474&action=grading", 'U17'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51403&action=grading", 'U18'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51332&action=grading", 'U19'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51261&action=grading", 'U20'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51190&action=grading", 'U21'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51119&action=grading", 'U22'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=51048&action=grading", 'U23'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50977&action=grading", 'U24'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50693&action=grading", 'U25'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50906&action=grading", 'U26'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50835&action=grading", 'U27'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50622&action=grading", 'U28'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50764&action=grading", 'U29'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50551&action=grading", 'U30'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50480&action=grading", 'U31'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50409&action=grading", 'U32'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50338&action=grading", 'U33'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50267&action=grading", 'U34'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50054&action=grading", 'U35'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50125&action=grading", 'U36'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=50196&action=grading", 'U37'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48393&action=grading", 'Z1'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48591&action=grading", 'Z2'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48690&action=grading", 'Z3'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48789&action=grading", 'Z4'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48492&action=grading", 'Z5'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48888&action=grading", 'Z6'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=48987&action=grading", 'Z7'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49086&action=grading", 'Z8'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49284&action=grading", 'Z9'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49185&action=grading", 'Z10'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49878&action=grading", 'Z11'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49779&action=grading", 'Z12'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49680&action=grading", 'Z13'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49977&action=grading", 'Z14'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49581&action=grading", 'Z15'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49383&action=grading", 'Z16'],
    ["https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=49482&action=grading", 'Z17']
]


def read_sprint(links, numeroSprint, driver, hoja_excel):
    for i in range(len(links)):
        contador = 1
        driver.get(links[i][0])
        flag = True

        while flag == True:
            try:

                codigo = driver.find_element(By.XPATH,
                                             f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[4]").text
                codigo_id = driver.find_element(By.XPATH,
                                                f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[5]").text
                nombre = driver.find_element(By.XPATH,
                                             f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[3]").text
                estado = driver.find_elements(By.XPATH,
                                              f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[6]/div")
                calificacion_final = driver.find_element(By.XPATH,
                                                         f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[4]/table/tbody/tr[{contador}]/td[15]").text
                # ws1 = wb.active

                for j in range(len(estado)):
                    if j == 0:
                        hoja_excel.append([codigo, codigo_id, nombre, estado[j].text, calificacion_final, links[i][1],
                                           f"s{numeroSprint}"])
                    else:
                        hoja_excel.append(["", "", "", estado[j].text, "", "", ""])

                contador += 1
            except:
                print(f"{links[i][1]} terminado")
                flag = False


read_sprint(links_sprint_1, 2, driver, ws1)
wb.save(dest_filename)

driver.quit()