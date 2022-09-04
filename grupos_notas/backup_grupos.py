import shutil
import os
from datetime import date
from datetime import datetime
import os


def backup():
    source_grupos = r'C:\Users\TOMASGONZALEZ\OneDrive - Universidad Industrial de Santander\MinTIC2_Ciclo3 - Archivos de MISION TIC - Coordinador de Proyectos\10in\11dbclt\1101grupos\Extraccion_grupos_dinamico'
    today = date.today()

    os.makedirs(r'C:\Users\TOMASGONZALEZ\Documents\backup\grupos' + f"\\{date.today()}",exist_ok= True)
    destination = r'C:\Users\TOMASGONZALEZ\Documents\backup\grupos' + f"\\{date.today()}"
    grupos = os.listdir(source_grupos)



    
    for file in grupos:
        new_path = shutil.move(f"{source_grupos}/{file}", destination)
        print(new_path)

if __name__ == "__main__":
    backup()