import shutil
import os
from datetime import date
from datetime import datetime
import os


def backup():
    source_notas = r'C:\Users\TOMASGONZALEZ\OneDrive - Universidad Industrial de Santander\MinTIC2_Ciclo3 - Archivos de MISION TIC - Coordinador de Proyectos\10in\11dbclt\1100notas\notas_dinamico'
    source_grupos = r'C:\Users\TOMASGONZALEZ\OneDrive - Universidad Industrial de Santander\MinTIC2_Ciclo3 - Archivos de MISION TIC - Coordinador de Proyectos\10in\11dbclt\1101grupos\Extraccion_grupos_dinamico'
    today = date.today()

    os.makedirs(r'C:\Users\TOMASGONZALEZ\Documents\backup' + f"\\{date.today()}")
    destination = r'C:\Users\TOMASGONZALEZ\Documents\backup' + f"\\{date.today()}"
    notas = os.listdir(source_notas)
    grupos = os.listdir(source_grupos)



    for file in notas:
        new_path = shutil.move(f"{source_notas}/{file}", destination)
        print(new_path)

    for file1 in grupos:
        new_path = shutil.move(f"{source_grupos}/{file1}", destination)
        print(new_path)

if __name__ == "__main__":
    backup()