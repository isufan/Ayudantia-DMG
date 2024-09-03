# Probando automatización de los emails

# Automatización de correos

# estos son packages a usar 
# sino están en el computador se deben instalar usando este comando

# installed.packages(c("emayili","rio","dplyr","stringr","readxl","tcltk","writexl"))

library(emayili)
library(rio)
library(dplyr)
library(stringr)
library(readxl)
library(writexl)
library(tcltk)
library(kableExtra)

# Primero cargar excels 

actuacion_rev_UA <- read_excel("bd_2024-1/UAS/Actuación_Solicitud cursos.xlsx",
                               sheet = "A revisar por UA")

col_var <- colnames(actuacion_rev_UA)

agronomia_rev_UA <- read_excel("bd_2024-1/UAS/Agronomía_Solicitud cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
antropologia_rev_UA <- read_excel("bd_2024-1/UAS/Antropología_Solicitud cursos.xlsx",
                                  sheet = "A revisar por UA",col_names = col_var)
arquitectura_rev_UA <- read_excel("bd_2024-1/UAS/Arquitectura_Solicitud cursos.xlsx",
                                  sheet = "A revisar por UA",col_names = col_var)
arte_rev_UA <- read_excel("bd_2024-1/UAS/Arte_Solicitud cursos.xlsx",
                          sheet = "A revisar por UA",col_names = col_var)
astrofisica_rev_UA <- read_excel("bd_2024-1/UAS/Astrofísica_Solicitud cursos.xlsx",
                                 sheet = "A revisar por UA",col_names = col_var)
ciencia_politica_rev_UA <- read_excel("bd_2024-1/UAS/Ciencia Política_Solicitud cursos.xlsx",
                                      sheet = "A revisar por UA",col_names = col_var)
ciencias_biologicas_rev_UA <- read_excel("bd_2024-1/UAS/Ciencias Biológicas_Solicitud cursos.xlsx",
                                         sheet = "A revisar por UA",col_names = col_var)
ciencias_salud_rev_UA <- read_excel("bd_2024-1/UAS/Ciencias de la Salud_Solicitud cursos.xlsx",
                                    sheet = "A revisar por UA",col_names = col_var)
comunicaciones_rev_UA <- read_excel("bd_2024-1/UAS/Comunicaciones_Solicitud cursos.xlsx",
                                    sheet = "A revisar por UA",col_names = col_var)
construccion_civil_rev_UA <- read_excel("bd_2024-1/UAS/Construcción Civil_Solicitud cursos.xlsx",
                                        sheet = "A revisar por UA",col_names = col_var)
derecho_rev_UA <- read_excel("bd_2024-1/UAS/Derecho_Solicitud cursos.xlsx",
                             sheet = "A revisar por UA",col_names = col_var)
des_sustentable_rev_UA <- read_excel("bd_2024-1/UAS/Desarrollo Sustentable_Solicitud cursos.xlsx",
                                     sheet = "A revisar por UA",col_names = col_var)
diseño_rev_UA <- read_excel("bd_2024-1/UAS/Diseño_Solicitud cursos.xlsx",
                            sheet = "A revisar por UA",col_names = col_var)
eco_admin_rev_UA <- read_excel("bd_2024-1/UAS/Economía y Administración_Solicitud cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
educacion_rev_UA <- read_excel("bd_2024-1/UAS/Educación_Solicitud cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
enfermeria_rev_UA <- read_excel("bd_2024-1/UAS/Enfermería_Solicitud cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
esc_gobierno_rev_UA <- read_excel("bd_2024-1/UAS/Escuela de Gobierno_Solicitud_cursos.xlsx",
                                  sheet = "A revisar por UA",col_names = col_var)
estetica_rev_UA <- read_excel("bd_2024-1/UAS/Estética_Solicitud_cursos.xlsx",
                              sheet = "A revisar por UA",col_names = col_var)
est_urb_rev_UA <- read_excel("bd_2024-1/UAS/Estudios Urbanos_Solicitud_cursos.xlsx",
                             sheet = "A revisar por UA",col_names = col_var)
filosofia_rev_UA <- read_excel("bd_2024-1/UAS/Filosofía_Solicitud_cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
fisica_rev_UA <- read_excel("bd_2024-1/UAS/Física_Solicitud_cursos.xlsx",
                            sheet = "A revisar por UA",col_names = col_var)
geografia_rev_UA <- read_excel("bd_2024-1/UAS/Geografía_Solicitud_cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
historia_rev_UA <- read_excel("bd_2024-1/UAS/Historia_Solicitud_cursos.xlsx",
                              sheet = "A revisar por UA",col_names = col_var)
ingenieria_rev_UA <- read_excel("bd_2024-1/UAS/Ingeniería_Solicitud_cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
letras_rev_UA <- read_excel("bd_2024-1/UAS/Letras_Solicitud_cursos.xlsx",
                            sheet = "A revisar por UA",col_names = col_var)
matematicas_rev_UA <- read_excel("bd_2024-1/UAS/Matemáticas_Solicitud_cursos.xlsx",
                                 sheet = "A revisar por UA",col_names = col_var)
medicina_rev_UA <- read_excel("bd_2024-1/UAS/Medicina_Solicitud_cursos.xlsx",
                              sheet = "A revisar por UA",col_names = col_var)
musica_rev_UA <- read_excel("bd_2024-1/UAS/Música_Solicitud_cursos.xlsx",
                            sheet = "A revisar por UA",col_names = col_var)
psicologia_rev_UA <- read_excel("bd_2024-1/UAS/Psicología_Solicitud_cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
quimica_rev_UA <- read_excel("bd_2024-1/UAS/Química_Solicitud_cursos.xlsx",
                             sheet = "A revisar por UA",col_names = col_var)
sociologia_rev_UA <- read_excel("bd_2024-1/UAS/Sociología_Solicitud_cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
teologia_rev_UA <- read_excel("bd_2024-1/UAS/Teología_Solicitud_cursos.xlsx",
                              sheet = "A revisar por UA",col_names = col_var)
trabajo_social_rev_UA <- read_excel("bd_2024-1/UAS/Trabajo Social_Solicitud_cursos.xlsx",
                                    sheet = "A revisar UA")

actuacion_sobre <- read_excel("bd_2024-1/UAS/Actuación_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG")

col_var2 <- colnames(actuacion_sobre)
agronomia_sobre <- read_excel("bd_2024-1/UAS/Agronomía_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
antropologia_sobre <- read_excel("bd_2024-1/UAS/Antropología_Solicitud cursos.xlsx",
                                 sheet = "Sobrepasos DMG",col_names = col_var2)
arquitectura_sobre <- read_excel("bd_2024-1/UAS/Arquitectura_Solicitud cursos.xlsx",
                                 sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
arte_sobre <- read_excel("bd_2024-1/UAS/Arte_Solicitud cursos.xlsx",
                         sheet = "Sobrepasos DMG",col_names = col_var2)
astrofisica_sobre <- read_excel("bd_2024-1/UAS/Astrofísica_Solicitud cursos.xlsx",
                                sheet = "Sobrepasos DMG",col_names = col_var2)
ciencia_politica_sobre <- read_excel("bd_2024-1/UAS/Ciencia Política_Solicitud cursos.xlsx",
                                     sheet = "Sobrepasos DMG",col_names = col_var2)
ciencias_biologicas_sobre <- read_excel("bd_2024-1/UAS/Ciencias Biológicas_Solicitud cursos.xlsx",
                                        sheet = "Sobrepasos DMG",col_names = col_var2)
ciencias_salud_sobre <- read_excel("bd_2024-1/UAS/Ciencias de la Salud_Solicitud cursos.xlsx",
                                   sheet = "Sobrepasos DMG",col_names = col_var2)
comunicaciones_sobre <- read_excel("bd_2024-1/UAS/Comunicaciones_Solicitud cursos.xlsx",
                                   sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
construccion_civil_sobre <- read_excel("bd_2024-1/UAS/Construcción Civil_Solicitud cursos.xlsx",
                                       sheet = "Sobrepasos DMG",col_names = col_var2)
derecho_sobre <- read_excel("bd_2024-1/UAS/Derecho_Solicitud cursos.xlsx",
                            sheet = "Sobrepasos DMG",col_names = col_var2)
des_sustentable_sobre <- read_excel("bd_2024-1/UAS/Desarrollo Sustentable_Solicitud cursos.xlsx",
                                    sheet = "Sobrepasos DMG",col_names = col_var2)
diseño_sobre <- read_excel("bd_2024-1/UAS/Diseño_Solicitud cursos.xlsx",
                           sheet = "Sobrepasos DMG",col_names = col_var2)
eco_admin_sobre <- read_excel("bd_2024-1/UAS/Economía y Administración_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
educacion_sobre <- read_excel("bd_2024-1/UAS/Educación_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
enfermeria_sobre <- read_excel("bd_2024-1/UAS/Enfermería_Solicitud cursos.xlsx",
                               sheet = "Sobrepasos DMG",col_names = col_var2)
esc_gobierno_sobre <- read_excel("bd_2024-1/UAS/Escuela de Gobierno_Solicitud_cursos.xlsx",
                                 sheet = "Sobrepasos DMG",col_names = col_var2)
estetica_sobre <- read_excel("bd_2024-1/UAS/Estética_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
est_urb_sobre <- read_excel("bd_2024-1/UAS/Estudios Urbanos_Solicitud_cursos.xlsx",
                            sheet = "Sobrepasos DMG",col_names = col_var2)
filosofia_sobre <- read_excel("bd_2024-1/UAS/Filosofía_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
fisica_sobre <- read_excel("bd_2024-1/UAS/Física_Solicitud_cursos.xlsx",
                           sheet = "Sobrepasos DMG",col_names = col_var2)
geografia_sobre <- read_excel("bd_2024-1/UAS/Geografía_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
historia_sobre <- read_excel("bd_2024-1/UAS/Historia_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
ingenieria_sobre <- read_excel("bd_2024-1/UAS/Ingeniería_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG",col_names = col_var2)
letras_sobre <- read_excel("bd_2024-1/UAS/Letras_Solicitud_cursos.xlsx",
                           sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
matematicas_sobre <- read_excel("bd_2024-1/UAS/Matemáticas_Solicitud_cursos.xlsx",
                                sheet = "Sobrepasos DMG",col_names = col_var2)
medicina_sobre <- read_excel("bd_2024-1/UAS/Medicina_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
musica_sobre <- read_excel("bd_2024-1/UAS/Música_Solicitud_cursos.xlsx",
                           sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
psicologia_sobre <- read_excel("bd_2024-1/UAS/Psicología_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG",col_names = col_var2)
quimica_sobre <- read_excel("bd_2024-1/UAS/Química_Solicitud_cursos.xlsx",
                            sheet = "Sobrepasos DMG",col_names = col_var2)
sociologia_sobre <- read_excel("bd_2024-1/UAS/Sociología_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG",col_names = col_var2)
teologia_sobre <- read_excel("bd_2024-1/UAS/Teología_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
trabajo_social_sobre <- read_excel("bd_2024-1/UAS/Trabajo Social_Solicitud_cursos.xlsx",
                                   sheet = "Sobrepasos DMG",col_names = col_var2)

# combinar todos los excel

consolidado_rev_UA <- rbind(actuacion_rev_UA,
                     agronomia_rev_UA,
                     antropologia_rev_UA,
                     arquitectura_rev_UA,
                     arte_rev_UA,
                     astrofisica_rev_UA,
                     ciencia_politica_rev_UA,
                     ciencias_biologicas_rev_UA,
                     ciencias_salud_rev_UA,
                     comunicaciones_rev_UA,
                     construccion_civil_rev_UA,
                     derecho_rev_UA,
                     des_sustentable_rev_UA,
                     diseño_rev_UA,
                     eco_admin_rev_UA,
                     educacion_rev_UA,
                     enfermeria_rev_UA,
                     esc_gobierno_rev_UA,
                     estetica_rev_UA,
                     est_urb_rev_UA,
                     filosofia_rev_UA,
                     fisica_rev_UA,
                     geografia_rev_UA,
                     historia_rev_UA,
                     ingenieria_rev_UA,
                     letras_rev_UA,
                     matematicas_rev_UA,
                     medicina_rev_UA,
                     musica_rev_UA,
                     psicologia_rev_UA,
                     quimica_rev_UA,
                     sociologia_rev_UA,
                     teologia_rev_UA,
                     trabajo_social_rev_UA
                     ) %>% 
  arrange("Nombre estudiante")

consolidado_sobre <- rbind(actuacion_sobre,
                            agronomia_sobre,
                            antropologia_sobre,
                            arquitectura_sobre,
                            arte_sobre,
                            astrofisica_sobre,
                            ciencia_politica_sobre,
                            ciencias_biologicas_sobre,
                            ciencias_salud_sobre,
                            comunicaciones_sobre,
                            construccion_civil_sobre,
                            derecho_sobre,
                            des_sustentable_sobre,
                            diseño_sobre,
                            eco_admin_sobre,
                            educacion_sobre,
                            enfermeria_sobre,
                            esc_gobierno_sobre,
                            estetica_sobre,
                            est_urb_sobre,
                            filosofia_sobre,
                            fisica_sobre,
                            geografia_sobre,
                            historia_sobre,
                            ingenieria_sobre,
                            letras_sobre,
                            matematicas_sobre,
                            medicina_sobre,
                            musica_sobre,
                            psicologia_sobre,
                            quimica_sobre,
                            sociologia_sobre,
                            teologia_sobre,
                            trabajo_social_sobre
) %>% 
  arrange(`Nombre estudiante`)

# seleccionamos el nombre de cada estudiante de forma única

estudiantes <- unique(consolidado_sobre$`Nombre estudiante`)

# estos es transversal para todos los correos

source("codigo_solicitar_contraseña.R")
usuario <- "fvilca@uc.cl"
contrasenya <- get_password()
asunto <- "[IMPORTANTE] Actualización solicitudes de cursos UC"

# comenzamos a escribir cada correo

for (i in estudiantes) {
  # filtramos el consolidado para el estudiante i
  cursos_aux <- filter(consolidado_sobre,`Nombre estudiante` == i)
  # se le da el correo del estudiante i 
  para <- cursos_aux$`Correo electrónico`[1]
  
  # creamos una tabla con la info de los cursos para el estudiante i
  tabla_sobre <- cursos_aux %>% 
    select(`Nombre curso`,Sigla,NRC,`Curso dentro del catálogo`) %>%
    kable(align = "c",booktabs = TRUE,
          col.names = c("Nombre Curso","Sigla", "NRC","Curso en Catálogo"))  %>%   
    row_spec(0, bold = T, color = "white", background = "#173F8A") %>%  
    kable_minimal()
  
  cursos_aux_UA <- filter(consolidado_rev_UA,`Nombre estudiante` == i)
  
  if (nrow(cursos_aux_UA) != 0) {
    tabla_rev_UA <- cursos_aux_UA %>% 
      select(`Nombre curso`,Sigla,NRC,`Respuesta UA`) %>%
      kable(align = "c",booktabs = TRUE,
            col.names = c("Nombre Curso","Sigla", "NRC","Respuesta UA"))  %>%   
      row_spec(0, bold = T, color = "white", background = "#173F8A") %>%  
      kable_minimal()
  }else{
    tabla_rev_UA <- "......"
  }
  
  # se crea el cuerpo del correo en formato HTML
  
  mensaje <- str_c(
    "
<p><strong><em>Este es un correo automático, no responder.</em></strong></p>
 
<p>Estimado/a estudiante,</p>
   
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
  
}
