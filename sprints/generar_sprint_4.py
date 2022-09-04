from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

from openpyxl import Workbook
from openpyxl.utils import get_column_letter



pathDriver = "C:/Users/TOMASGONZALEZ/chromedriver.exe"
usuario=1007414252
contrasena="MisionTIC2022$"

wb = Workbook()
dest_filename = 'sprint_1.xlsx'
ws1 = wb.active
ws1.title = "sprint_1"
ws1.append(["Codigo","Codigo_Id","Nombre_Tripulante", "Estado","Nota_Final","Grupo","Sprint"])

chrome_options = Options()
#chrome_options.add_argument("--disable-extensions")
#chrome_options.add_argument("--disable-gpu")
#chrome_options.add_argument("--headless")
chrome_options.headless = True
driver = webdriver.Chrome(executable_path=pathDriver,options=chrome_options)
driver.implicitly_wait(0.5)


driver.get("https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37842&action=grading") #O1
username=driver.find_element(By.ID,"username")
password=driver.find_element(By.ID,"password")
login = driver.find_element(By.ID,"loginbtn")
username.send_keys(usuario)
password.send_keys(contrasena)
login.click()

links_sprint_4 = [
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37903&action=grading",#O1
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38089&action=grading",#O2
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=37996&action=grading",#O3
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38275&action=grading",#O4
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38182&action=grading",#O5
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38349&action=grading",#O6
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38554&action=grading",#O7
    "",#O8
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38461&action=grading",#O9
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38647&action=grading",#O10
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39391&action=grading",#O11
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39484&action=grading",#O12
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39205&action=grading",#O13
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39298&action=grading",#O14
    "",#O15
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39019&action=grading",#O16
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39112&action=grading",#O17
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=38833&action=grading",#O18
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39577&action=grading",#O19
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39763&action=grading",#O20
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39856&action=grading",#O21
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39949&action=grading",#O22
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40042&action=grading",#O23
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=39670&action=grading",#O24
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40321&action=grading",#O25
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40414&action=grading",#O26
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40135&action=grading",#O27
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40228&action=grading",#O28
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40507&action=grading",#O29
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40600&action=grading",#O30
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40879&action=grading",#O31
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40693&action=grading",#O32
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41158&action=grading",#O33
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40786&action=grading",#O34
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41251&action=grading",#O35
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=40972&action=grading",#O36
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41065&action=grading",#O37
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41344&action=grading",#O38
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41530&action=grading",#O39
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41716&action=grading",#O40
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41623&action=grading",#O41
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41809&action=grading",#O42
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42088&action=grading",#O43
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42274&action=grading",#O44
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41995&action=grading",#O45
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41902&action=grading",#O46
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=41437&action=grading",#O47
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42181&action=grading",#O48
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42460&action=grading",#O49
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42367&action=grading",#O50
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42553&action=grading",#O51
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42646&action=grading",#O52
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42739&action=grading",#O53
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42925&action=grading",#O54
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=42832&action=grading",#O55
    "",#O56
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43111&action=grading",#O57
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43204&action=grading",#O58
    "https://lms.uis.edu.co/mintic2022/mod/assign/view.php?id=43297&action=grading",#O59

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
                nombre = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[1]").text
                estado = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[6]").text
                calificacion_final = driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/div[5]/div[3]/div[2]/div/div/div/div/div/div[3]/div[{a}]/table/tbody/tr[{contador}]/td[15]").text
                #ws1 = wb.active
                hoja_excel.append([codigo,codigo_id,nombre,estado,calificacion_final,f"O{i+1}",f"s{numeroSprint}"])
                contador+=1
            except:
                print(f"O{i+1} terminado")
                flag = False
        
        

read_sprint(links_sprint_4,4,driver,ws1)
wb.save(dest_filename)

driver.quit()
