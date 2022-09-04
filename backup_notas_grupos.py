import time
inicio = time.time()
#exec(open("./grupos_notas/backup_grupos.py").read())
#exec(open("./grupos_notas/backup_notas.py").read())
#exec(open("generar_sprint_1.py").read())
#exec(open("generar_sprint_2.py").read())
exec(open("./grupos_notas/descargar_notas.py").read())
exec(open("./grupos_notas/descargar_grupos.py").read())
fin = time.time()
print(fin-inicio)


