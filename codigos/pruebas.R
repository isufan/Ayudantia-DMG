actuacion_sobre <- read_excel("bd_2024-1/UAS/Actuación_Solicitud cursos.xlsx",
                               sheet = "Sobrepasos DMG")
agronomia_sobre <- read_excel("bd_2024-1/UAS/Agronomía_Solicitud cursos.xlsx",
                               sheet = "Sobrepasos DMG")
antropologia_sobre <- read_excel("bd_2024-1/UAS/Antropología_Solicitud cursos.xlsx",
                                  sheet = "Sobrepasos DMG")
arquitectura_sobre <- read_excel("bd_2024-1/UAS/Arquitectura_Solicitud cursos.xlsx",
                                  sheet = "Sobrepasos DMG")
arte_sobre <- read_excel("bd_2024-1/UAS/Arte_Solicitud cursos.xlsx",
                          sheet = "Sobrepasos DMG")
astrofisica_sobre <- read_excel("bd_2024-1/UAS/Astrofísica_Solicitud cursos.xlsx",
                                 sheet = "Sobrepasos DMG")
ciencia_politica_sobre <- read_excel("bd_2024-1/UAS/Ciencia Política_Solicitud cursos.xlsx",
                                      sheet = "Sobrepasos DMG")
ciencias_biologicas_sobre <- read_excel("bd_2024-1/UAS/Ciencias Biológicas_Solicitud cursos.xlsx",
                                         sheet = "Sobrepasos DMG")
ciencias_salud_sobre <- read_excel("bd_2024-1/UAS/Ciencias de la Salud_Solicitud cursos.xlsx",
                                    sheet = "Sobrepasos DMG")
comunicaciones_sobre <- read_excel("bd_2024-1/UAS/Comunicaciones_Solicitud cursos.xlsx",
                                    sheet = "Sobrepasos DMG")
construccion_civil_sobre <- read_excel("bd_2024-1/UAS/Construcción Civil_Solicitud cursos.xlsx",
                                        sheet = "Sobrepasos DMG")
derecho_sobre <- read_excel("bd_2024-1/UAS/Derecho_Solicitud cursos.xlsx",
                             sheet = "Sobrepasos DMG")
des_sustentable_sobre <- read_excel("bd_2024-1/UAS/Desarrollo Sustentable_Solicitud cursos.xlsx",
                                     sheet = "Sobrepasos DMG")
diseño_sobre <- read_excel("bd_2024-1/UAS/Diseño_Solicitud cursos.xlsx",
                            sheet = "Sobrepasos DMG")
eco_admin_sobre <- read_excel("bd_2024-1/UAS/Economía y Administración_Solicitud cursos.xlsx",
                               sheet = "Sobrepasos DMG")
educacion_sobre <- read_excel("bd_2024-1/UAS/Educación_Solicitud cursos.xlsx",
                               sheet = "Sobrepasos DMG")
enfermeria_sobre <- read_excel("bd_2024-1/UAS/Enfermería_Solicitud cursos.xlsx",
                                sheet = "Sobrepasos DMG")
esc_gobierno_sobre <- read_excel("bd_2024-1/UAS/Escuela de Gobierno_Solicitud_cursos.xlsx",
                                  sheet = "Sobrepasos DMG")
estetica_sobre <- read_excel("bd_2024-1/UAS/Estética_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG")
est_urb_sobre <- read_excel("bd_2024-1/UAS/Estudios Urbanos_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG")
filosofia_sobre <- read_excel("bd_2024-1/UAS/Filosofía_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG")
fisica_sobre <- read_excel("bd_2024-1/UAS/Física_Solicitud_cursos.xlsx",
                            sheet = "Sobrepasos DMG")
geografia_sobre <- read_excel("bd_2024-1/UAS/Geografía_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG")
historia_sobre <- read_excel("bd_2024-1/UAS/Historia_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG")
ingenieria_sobre <- read_excel("bd_2024-1/UAS/Ingeniería_Solicitud_cursos.xlsx",
                                sheet = "Sobrepasos DMG")
letras_sobre <- read_excel("bd_2024-1/UAS/Letras_Solicitud_cursos.xlsx",
                            sheet = "Sobrepasos DMG")
matematicas_sobre <- read_excel("bd_2024-1/UAS/Matemáticas_Solicitud_cursos.xlsx",
                                 sheet = "Sobrepasos DMG")
medicina_sobre <- read_excel("bd_2024-1/UAS/Medicina_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG")
musica_sobre <- read_excel("bd_2024-1/UAS/Música_Solicitud_cursos.xlsx",
                            sheet = "Sobrepasos DMG")
psicologia_sobre <- read_excel("bd_2024-1/UAS/Psicología_Solicitud_cursos.xlsx",
                                sheet = "Sobrepasos DMG")
quimica_sobre <- read_excel("bd_2024-1/UAS/Química_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG")
sociologia_sobre <- read_excel("bd_2024-1/UAS/Sociología_Solicitud_cursos.xlsx",
                                sheet = "Sobrepasos DMG")
teologia_sobre <- read_excel("bd_2024-1/UAS/Teología_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG")
trabajo_social_sobre <- read_excel("bd_2024-1/UAS/Trabajo Social_Solicitud_cursos.xlsx",
                                    sheet = "Sobrepasos DMG")
names(consolidado_sobre2) == names(consolidado_rev_UA)
names(arte_rev_UA)
names(ciencias_salud_rev_UA)
names(comunicaciones_rev_UA)
names(construccion_civil_rev_UA)
names(derecho_rev_UA)
names(des_sustentable_rev_UA)
names(diseño_rev_UA)
names(eco_admin_rev_UA)
names(educacion_rev_UA)
names(enfermeria_rev_UA)
names(esc_gobierno_rev_UA)
names(estetica_rev_UA)
names(est_urb_rev_UA)
names(filosofia_rev_UA)
names(fisica_rev_UA)
names(geografia_rev_UA)
names(historia_rev_UA)
names(ingenieria_rev_UA)
names(letras_rev_UA)
names(matematicas_rev_UA)
names(medicina_rev_UA)
names(musica_rev_UA)
names(psicologia_rev_UA)
names(quimica_rev_UA)
names(sociologia_rev_UA)
names(teologia_rev_UA)
names(trabajo_social_rev_UA)
names####

dim(consolidado_sobre)
dim(agronomia_rev_UA)
dim(astrofisica_rev_UA)
dim(ciencia_politica_rev_UA)
dim(ciencias_biologicas_rev_UA)

names(consolidado_sobre2)
names(consolidado_rev_UA)

consolidado_sobre2 <- consolidado_sobre %>% 
  mutate("Respuesta UA" = rep("NO APLICA",2191)) %>% 
  mutate("Observación UA" = rep("NO APLICA",2191)) %>% 
  select(c(1:15),c(18:19),c(17,16)) %>% 
  rename("Fecha de solicitud del estudiante" = `Fecha de solicitud`)

consolidado_final = rbind(consolidado_sobre2,consolidado_rev_UA) %>% 
  mutate(Categoria = c(rep("sobre",2191),rep("UA",290)) ) %>% 
  filter(`Nombre estudiante` != "Nombre completo" & `Nombre estudiante` != "NA" )

aaaa <- unique(consolidado_final$`Nombre estudiante`)

aux <- c()

for (i in aaaa) {
  if (!(i %in% estudiantes)  ) {
    aux <- c(aux,i)
  }
}

estudiantes_pendientes <- consolidado_final %>% 
  filter(`Nombre estudiante` %in% aux)


write_xlsx(estudiantes_pendientes, "resultados/pendientes_correo.xlsx")

cursos_aux <- filter(estudiantes_pendientes,`Nombre estudiante` == i & Categoria == "sobre")

# creamos una tabla con la info de los cursos para el estudiante i
if (nrow(cursos_aux) != 0) {
  # se le da el correo del estudiante i 
  para <- cursos_aux$`Correo electrónico`[1]
  tabla_sobre <- cursos_aux %>% 
    select(`Nombre curso`,Sigla,NRC,`Curso dentro del catálogo`) %>%
    kable(align = "c",booktabs = TRUE,
          col.names = c("Nombre Curso","Sigla", "NRC","Curso en Catálogo"))  %>%   
    row_spec(0, bold = T, color = "white", background = "#173F8A") %>%  
    kable_minimal()
}else{
  tabla_sobre <- "......"
}

cursos_aux_UA <- filter(estudiantes_pendientes,`Nombre estudiante` == i & Categoria == "UA")

if (nrow(cursos_aux_UA) != 0) {
  # se le da el correo del estudiante i 
  para <- cursos_aux_UA$`Correo electrónico`[1]
  tabla_rev_UA <- cursos_aux_UA %>% 
    select(`Nombre curso`,Sigla,NRC,`Respuesta UA`) %>%
    kable(align = "c",booktabs = TRUE,
          col.names = c("Nombre Curso","Sigla", "NRC","Respuesta UA"))  %>%   
    row_spec(0, bold = T, color = "white", background = "#173F8A") %>%  
    kable_minimal()
}else{
  tabla_rev_UA <- "......"
}


mensaje <- str_c(
  "
<p><strong><em>Este es un correo automático, no responder.</em></strong></p>
 
<p>Estimados y estimadas estudiantes,</p>
   
<p>Junto con saludar y esperando que te encuentres muy bien, escribimos este correo para entregarte información sobre el proceso de inscripción de cursos para este semestre.</p>
 
<h3>Cursos Inscritos</h3>
<p>Basados en tu preselección, te inscribimos en uno o dos cursos en tu carga académica. Puedes revisarlo accediendo al Portal UC con tu usuario y contraseña UC, en la pestaña “información académica”.</p>
 
<h3>Periodo de inscripción por parte cada estudiante</h3>
<p>La plataforma Banner UC estará disponible para complementar o modificar tu carga académica desde mañana <strong>31 de julio hasta el miércoles 14 de agosto</strong>, de 08:00 a 19:50 horas (hora chilena) (de lunes a viernes).</p>
 
<h3>Cursos para inscripción a través de Banner UC</h3>
<p>Los cursos que solicitaste que <strong>sí cuentan con vacantes reservadas</strong> para estudiantes de intercambio, estarán habilitados para que los inscribas desde mañana 31 de julio a través de la plataforma Banner UC:</p>
",
tabla_sobre,

"<p>Si solicitaste cursos que <strong>no cuentan con vacantes reservadas</strong> para estudiantes de intercambio, comentarte que estos fueron enviados a las respectivas Unidades Académicas de la Universidad para que evaluaran la solicitud de una vacante adicional. Las respuestas de cada curso significan:</p>
  <ul>
  <li><strong>“Sin revisar”:</strong> hasta el momento la Unidad Académica no ha revisado la solicitud.</li>
  <li><strong>“En revisión”:</strong> la Unidad Académica está revisando la solicitud.</li>
  <li><strong>“Autorizado”:</strong> la Unidad Académica sí autorizó una vacante adicional, por lo que podrás inscribir el curso. Si al intentar inscribir el curso en el sistema aparece “error”, se debe a que aún se está gestionando el reajuste de la vacante en el sistema. Debes seguir intentando inscribir el curso durante los próximos días.</li>
  <li><strong>“No autorizado”:</strong> la Unidad Académica no autorizó una vacante adicional, por lo que no podrás inscribir el curso.</li>
  </ul>",
tabla_rev_UA,
"<h3>Acceso a la plataforma</h3>
  <p>Para acceder a la plataforma sigue esta ruta:</p>
  <p> <a href='https://portal.uc.cl/'>Portal.uc.cl</a> > <strong>
  Información Académica > Herramientas del estudiante > Agregar o eliminar cursos </strong> </p>
  
  <h3>Video Tutorial</h3>
  <p>Revisa este video tutorial en el siguiente <a href='https://www.youtube.com/watch?v=sbeU-LNyuZA'> enlace </a> sobre cómo usar la plataforma Banner, enfocándote especialmente en el segmento del minuto <strong>1.34 al 3.30</strong> para revisar la información correspondiente a estudiantes de intercambio.</p>
  
  <h3>Información Importante:</h3>
  <ul>
  <li>Que el curso esté habilitado, no significa que esté asegurado, dado que las vacantes se asignan por orden de inscripción en la plataforma Banner. Sé proactivo/a.</li>
  <li>Si tienes un curso <strong>“no autorizado”</strong>, puede deberse a distintas situaciones como falta de espacio en sala, requisitos previos no cumplidos, entre otros.</li>
  <li>Si tienes un <strong>“curso no se dicta este semestre”</strong>, corresponde a cursos que no fueron parte de la programación académica 2024-2 o cursos que estuvieron inicialmente en el catálogo pero por motivos de fuerza mayor ya no se realizarán este semestre.</li>
  <li>Para aquellos que tengan cursos “sin revisar” o “en revisión”, se realizarán actualizaciones del estado de estos cursos los días miércoles 7, viernes 9, lunes 12, martes 13 y miércoles 14 de agosto.</li>
  <li>Aquellos estudiantes que hayan solicitado cursos deportivos serán contactados durante los próximos días.</li>
  <li>Los cursos de <a href='https://internacionalizacion.uc.cl/quiero-ir-a-la-uc/viajar-a-la-uc/aprende-espanol-en-uc-chile/cursos-de-espanol-uc/'> Español Lengua Extranjera </a> son cursos pagados y si desean cursarlos deben escribir a <a href='mailto:cbonillam@uc.cl'> cbonillam@uc.cl </a> para su inscripción.</li>
  </ul>
  
  <h3>Atención Presencial</h3>
  <p>Para dudas o situaciones especiales asociadas a la inscripción de cursos, debes dirigirte a las oficinas VRAI de Campus San Joaquín. La atención de estos casos será <strong>exclusivamente presencial</strong> y por orden de llegada.</p>
  
  <p><u>Dirección:</u> Av. Vicuña Mackenna 4860, Macul, Hall central, segundo piso.</p>
  <p><u>Horario de atención:</u> 09:30 a 12:30 horas</p>
  <p><u>Período de atención:</u> 5 al 14 de agosto.</p>
  <p><u><strong>No se atenderá en Campus Casa Central durante estos días.</strong></u></p>
  
  <p>Para situaciones excepcionales que te impidan visitar el campus, por favor escribe directamente a tu Coordinadora de acompañamiento (M. Belén Saavedra, Magdalena Cartagena, Sophia Acevedo o Úrsula Godoy).</p>
  
  <p>¡Estamos aquí para ayudarte a tener un excelente inicio de semestre!</p>",
"<p>Saludos cordiales,</p>",
"<p><strong>Coordinación de Acompañamiento</strong><br>",
"<em>Dirección de Movilidad Global</em><br>",
"<em>Vicerrectoría de Asuntos Internacionales</em></p>"
)


for (i in estudiantes) {
  # filtramos el consolidado para el estudiante i
  cursos_aux <- filter(estudiantes_pendientes,`Nombre estudiante` == estudiantes[3] & Categoria == "sobre")
  
  # creamos una tabla con la info de los cursos para el estudiante i
  if (nrow(cursos_aux) != 0) {
    # se le da el correo del estudiante i 
    para <- "fvilca@uc.cl"
    tabla_sobre <- cursos_aux %>% 
      select(`Fecha de solicitud del estudiante`,
             `Nombre curso`,Sigla,NRC,`Curso dentro del catálogo`) %>%
      kable(align = "c",booktabs = TRUE,
            col.names = c("Fecha solicitud","Nombre Curso","Sigla", "NRC","Curso en Catálogo"))  %>%   
      row_spec(0, bold = T, color = "white", background = "#173F8A") %>%  
      kable_minimal()
  }else{
    tabla_sobre <- "......"
  }
  
  cursos_aux_UA <- filter(estudiantes_pendientes,`Nombre estudiante` == estudiantes[3] & Categoria == "UA")
  
  if (nrow(cursos_aux_UA) != 0) {
    # se le da el correo del estudiante i 
    para <- "fvilca@uc.cl"
    tabla_rev_UA <- cursos_aux_UA %>% 
      select(`Fecha de solicitud del estudiante`,
             `Nombre curso`,NRC,Sigla,`Respuesta UA`) %>%
      kable(align = "c",booktabs = TRUE,
            col.names = c("Fecha solicitud","Nombre Curso","NRC","Sigla","Estado Solicitud"))  %>%   
      row_spec(0, bold = T, color = "white", background = "#173F8A") %>%  
      kable_minimal()
  }else{
    tabla_rev_UA <- "......"
  }
  
  # se crea el cuerpo del correo en formato HTML
  
  mensaje <- str_c(
    "
<p><strong><em>Este es un correo automático, no responder.</em></strong></p>
 
<p>Estimados y estimadas estudiantes,</p>
   
<p>Junto con saludar y esperando que se encuentren muy bien, les enviamos a continuación la actualización de los cursos solicitados:</p>·",
tabla_rev_UA,
"<h3><u>Importante:</u></h3>
<p>Los estados de la solicitud de cada curso significan:</p>
  <ul>
  <li><strong>“Sin revisar”:</strong> hasta el momento la Unidad Académica no ha revisado la solicitud. No es necesario que vuelvas a ir a nuestra oficina, debes esperar el correo de actualización del estado de las solicitudes de los cursos que enviaremos el día viernes 9 de agosto después de las 16:00 hrs.</li>
  <li><strong>“En revisión”:</strong> la Unidad Académica está revisando la solicitud. No es necesario que vuelvas a ir a la oficina, debes esperar el correo de actualización del estado de las solicitudes de los cursos que enviaremos el día viernes 9 de agosto después de las 16:00 hrs.</li>
  <li><strong>“Autorizado”:</strong> lla Unidad Académica sí autorizó una vacante adicional, por lo que podrás inscribir el curso. Si al intentar inscribir el curso en el sistema aparece “error”, se debe a que aún se está gestionando el reajuste de la vacante en el sistema. Te agradecemos lo inscribas durante los próximos días. De todas maneras, debes comenzar a asistir a las clases según el horario disponible que puedes encontrar en <a href='https://buscacursos.uc.cl/  '>buscacurso.uc.cl</a></li>
  <li><strong>“No autorizado”:</strong> la Unidad Académica no autorizó una vacante adicional, por lo que no podrás inscribir el curso. La no autorización de un curso puede deberse a distintas situaciones como falta de espacio en sala, requisitos previos no cumplidos, entre otros. </li>
  </ul>
<p>Por último, les recordamos que si necesitan apoyo para inscribir algún curso que no aparezca en el listado de arriba, pueden acercarse las oficinas VRAI de Campus San Joaquín (2° Piso Hall Universitario).</p>
  <p><u>Horario de atención:</u> 09:30 a 12:30 horas</p>
  <p><u>Período de atención:</u> 5 al 14 de agosto.</p>
<p>Que tengan un excelente día.</p>",
"<p>Saludos cordiales,</p>",
"<p><strong>Coordinación de Acompañamiento</strong><br>",
"<em>Dirección de Movilidad Global</em><br>",
"<em>Vicerrectoría de Asuntos Internacionales</em></p>"
  )
  # Aquí configuramos el servidor SMTP de outlook
  smtp <- server(host = "smtp.office365.com",
                 port = 587,
                 username = usuario,
                 password = contrasenya)
  
  # Aquí creamos el mensaje de correo electrónico y se le da el formato HTML
  email <- envelope() %>%
    from(usuario) %>%
    to(para) %>%
    subject(asunto) %>%
    html(mensaje)
  
  # Este es el comando que envía el correo 
  smtp(email)
  print(i)
}
write_xlsx(consolidado_rev_UA, 
           "resultados/consolidado_rev_UA.xlsx")

table(consolidado_rev_UA$`Respuesta UA`)
