import pandas as pd


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


def enviar_correo(mensaje, asunto, contrasena, emisor, receptor):

    mensaje = mensaje
    password = contrasena

    msg = MIMEMultipart()
    msg['From'] = emisor
    msg['Subject'] = asunto
    msg.attach(MIMEText(mensaje, 'html'))
    server = smtplib.SMTP('smtp.gmail.com: 587')
    server.starttls()
    server.login(msg['From'], password)
    msg['To'] = receptor
    server.sendmail(msg['From'], msg['To'], msg.as_string())


    server.quit()

    print
    "successfully sent email to %s:" % (msg['To'])



def main1():
    ORIGEN = 'pmomisiontic@gmail.com'
    DESTINO = 'tomasggrlol@gmail.com'
    CORREO_TUTOR = ""
    ASUNTO = 'Reporte de conformación de equipos y notas pendientes por calificación para formador'
    CONTRASENA = "gpyosptdqfidkyil"
    fecha = "29 de agosto de 2022"



    data = pd.read_excel("Tabla Estado Sprints.xlsx")
    data.columns = [column.replace(" ", "_") for column in data.columns]
    frame = pd.DataFrame(data)

    formadores = [
        "ALEX DAVID MORALES ACOSTA",
        "CARLOS ADOLFO BELTRAN CASTRO",
        "CARLOS ANDRES GUERRERO ALARCON",
        "CARLOS ANDRES PALMA SUAREZ",
        "CARLOS ARTURO PARRA ORTEGA",
        "CARLOS EDUARDO CLAROS ROZO",
        "CARLOS HORACIO GARCIA CASTRILLON",
        "CARLOS HUMBERTO CARRENO DIAZ",
        "CARLOS SANTIAGO BENITEZ MALAVER",
        "DANIEL RINCON LEON",
        "DARIO ALEJANDRO RIANO VELANDIA",
        "DAVID EDUARDO FLOREZ GOMEZ",
        "DIEGO FERNANDO MANTILLA VEGA",
        "EDDINSON JULIAN CARRILLO SISA",
        "EDILBERTO SIERRA MEJIA",
        "EDWARD ALBERTO ROPERO PEREZ",
        "GUILLERMO BEJARANO REYES",
        "JAHIR AUGUSTO SAAVEDRA GARCIA",
        "JAVIER ENRIQUE QUINTERO ROJAS",
        "JHON JAIRO CORTES PAREDES",
        "JORGE ARMANDO RODRIGUEZ HERNANDEZ",
        "LUIS ALBERTO SAURITH ALVARADO",
        "LUIS CARLOS GAMBOA GELVEZ",
        "OMAR VIVAS CALDERON",
        "OSCAR MAURICIO GONZALEZ GOMEZ",
        "ROGERIO ORLANDO BELTRAN CASTRO",
        "ROY HERNANDO LLAMAS MUNOZ",
        "SERGIO ARTURO MEDINA CASTILLO",
        "SERGIO NICOLAS PINILLA CANON",
        "SILVIA JULIANA  MORENO ROA",
        "YESID OSWALDO QUINTERO MARTINEZ",
        "YHARY ESTEFANIA ARIAS TRILLOS",
    ]

    for i in range(len(formadores)):

        for row in pd.DataFrame(frame.query(f'Formador == "{formadores[i]}"')).itertuples():
            CORREO_TUTOR = row.Email_Formador
            break


        print(CORREO_TUTOR)
        frame_pendientes_por_calificar_s1 = pd.DataFrame(frame.query(f'Formador == "{formadores[i]}" and Estado_S1 == "Pendiente por calificar"')).sort_values("Curso", ascending=True)
        frame_pendientes_por_calificar_s2 = pd.DataFrame(frame.query(f'Formador == "{formadores[i]}" and Estado_S2 == "Pendiente por calificar"')).sort_values("Curso", ascending=True)
        frame_pendientes_por_grupo = pd.DataFrame(frame.query(f'Formador == "{formadores[i]}" and Tipo_proyecto == "Sin contestar aún"')).sort_values("Curso", ascending=True)

        tabla_pendientes_por_calificacion_s1 = """
        <table>
          <tr>
            <th>Grupo</th>
            <th>Estudiante</th>
          </tr>
          """
        for row in frame_pendientes_por_calificar_s1.itertuples():
            tabla_pendientes_por_calificacion_s1 = tabla_pendientes_por_calificacion_s1 + f"""
            \b<tr>
              \b<td>{row.Curso}</td>
              \b<td>{row.Nombre_Tripulante}</td>
            \b</tr>
            """

        if frame_pendientes_por_calificar_s1.empty:
            tabla_pendientes_por_calificacion_s1 = "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;No quedan tripulantes pendientes por calificar</p>"
        else:
            tabla_pendientes_por_calificacion_s1 = tabla_pendientes_por_calificacion_s1 +  "\b\b\b</table>"

        tabla_pendientes_por_calificacion_s2 = """
                <table>
                  <tr>
                    <th>Grupo</th>
                    <th>Estudiante</th>
                  </tr>
                  """
        for row in frame_pendientes_por_calificar_s2.itertuples():
            tabla_pendientes_por_calificacion_s2 = tabla_pendientes_por_calificacion_s2 + f"""
                    \b<tr>
                      \b<td>{row.Curso}</td>
                      \b<td>{row.Nombre_Tripulante}</td>
                    \b</tr>
                    """

        if frame_pendientes_por_calificar_s2.empty:
            tabla_pendientes_por_calificacion_s2 = "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;No quedan tripulantes pendientes por calificar</p>"
        else:
            tabla_pendientes_por_calificacion_s2 = tabla_pendientes_por_calificacion_s2 + "\b\b\b</table>"

        tabla_pendientes_por_grupo = """
                <table>
                  <tr>
                    <th>Grupo</th>
                    <th>Estudiante</th>
                  </tr>
                  """

        for row in frame_pendientes_por_grupo.itertuples():
            tabla_pendientes_por_grupo = tabla_pendientes_por_grupo + f"""
                    \b<tr>
                      \b<td>{row.Curso}</td>
                      \b<td>{row.Nombre_Tripulante}</td>
                    \b</tr>
                    """

        if frame_pendientes_por_grupo.empty:
            tabla_pendientes_por_grupo = "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;No quedan tripulantes pendientes por asignar grupo</p>"
        else:
            tabla_pendientes_por_grupo = tabla_pendientes_por_grupo + "\b\b\b</table>"

        #formato
        style = """
      <style type="text/css">
      @import url('https://fonts.googleapis.com/css2?family=Roboto&display=swap');
      body{font-family: "Roboto";}
      p {color: black}
      .verde {color: green; margin: 0px}
      table{border-collapse: collapse; margin-left: 20px;}
      th{border: solid 1px; padding: 5px; text-align: center; background-color: rgb(194, 194, 255);}
      td{border: solid 1px; padding: 5px; text-align: center;}   
      </style>
      """

        content = f'''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    {style}
</head>
<body>
    <br>
    <p>Buenos días <strong>{formadores[i]}</strong></p>
    <pre></pre>
    <p>A continuación, se presenta la relación de entregas pendientes por calificar en el Moodle (parte 1) y conformación de grupos (parte 2) con corte al {fecha}.</p>
    <pre></pre>
    <p><strong>PARTE 1. REPORTE DE CALIFICACIONES PENDIENTES</strong></p>
    <pre></pre>

    
    
    <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>SPRINT 1</strong></p>
    
    
    {tabla_pendientes_por_calificacion_s1}
    
    <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>SPRINT 2</strong></p>
    
    
    {tabla_pendientes_por_calificacion_s2}

    <p><strong>PARTE 2. REPORTE DE TRIPULANTES PENDIENTES DE CONFORMACIÓN DE GRUPOS</strong></p>
    
    {tabla_pendientes_por_grupo}


    <p>De forma atenta, en caso de ser necesario, le solicitamos poner al día la calificación de las notas pendientes y la conformación de equipos. Adicionalmente agradecemos sus comentarios al reporte presentado, respondiendo este correo a los correos misiontic.prof2@uis.edu.co y misiontic.prof3@uis.edu.co.</p>
    <br>
    <p>Cordialmente,</p>
    <br>
    <p class="verde"><strong>EQUIPO PMO</strong></p>
    <p class="verde" >Profesionales Monitoreo de Proyectos Ciclo 3</p>
    <p class="verde"><a href=""></a> misiontic.prof2@uis.edu.co y misiontic.prof3@uis.edu.co</p>
    <p class="verde">3005172282 – 3208141002</p>
  </body>
</html>
'''

        enviar_correo(content,ASUNTO, CONTRASENA,ORIGEN,DESTINO)
        #enviar_correo(content,ASUNTO, CONTRASENA,ORIGEN,CORREO_TUTOR)
        #enviar_correo(content, ASUNTO, CONTRASENA, ORIGEN, "misiontic.prof2@uis.edu.co")





if __name__ == '__main__':
    main1()
