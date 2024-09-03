
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

actuacion_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Actuación_Solicitud cursos.xlsx",
                               sheet = "A revisar por UA")

col_var <- colnames(actuacion_rev_UA)

agronomia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Agronomía_Solicitud cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
antropologia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Antropología_Solicitud cursos.xlsx",
                                  sheet = "A revisar por UA",col_names = col_var)
arquitectura_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Arquitectura_Solicitud cursos.xlsx",
                                  sheet = "A revisar por UA",col_names = col_var)
arte_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Arte_Solicitud cursos.xlsx",
                          sheet = "A revisar por UA")%>% 
  select(c(1:19))
astrofisica_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Astrofísica_Solicitud cursos.xlsx",
                                 sheet = "A revisar por UA",col_names = col_var)
ciencia_politica_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Ciencia Política_Solicitud cursos.xlsx",
                                      sheet = "A revisar por UA",col_names = col_var)
ciencias_biologicas_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Ciencias Biológicas_Solicitud cursos.xlsx",
                                         sheet = "A revisar por UA",col_names = col_var)
ciencias_salud_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Ciencias de la Salud_Solicitud cursos.xlsx",
                                    sheet = "A revisar por UA",col_names = col_var)
comunicaciones_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Comunicaciones_Solicitud cursos.xlsx",
                                    sheet = "A revisar por UA",col_names = col_var)
construccion_civil_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Construcción Civil_Solicitud cursos.xlsx",
                                        sheet = "A revisar por UA",col_names = col_var)
derecho_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Derecho_Solicitud cursos (4).xlsx",
                             sheet = "A revisar por UA",col_names = col_var)
des_sustentable_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Desarrollo Sustentable_Solicitud cursos.xlsx",
                                     sheet = "A revisar por UA",col_names = col_var)
diseño_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Diseño_Solicitud cursos.xlsx",
                            sheet = "A revisar por UA",col_names = col_var)
eco_admin_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Economía y Administración_Solicitud cursos (4) 1.xlsx",
                               sheet = "A revisar por UA") %>% 
  select(c(1:19))
educacion_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Educación_Solicitud cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
enfermeria_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Enfermería_Solicitud cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
esc_gobierno_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Escuela de Gobierno_Solicitud_cursos.xlsx",
                                  sheet = "A revisar por UA",col_names = col_var)
estetica_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Estética_Solicitud_cursos.xlsx",
                              sheet = "A revisar por UA",col_names = col_var)
est_urb_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Estudios Urbanos_Solicitud_cursos.xlsx",
                             sheet = "A revisar por UA",col_names = col_var)
filosofia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Filosofía_Solicitud_cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
fisica_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Física_Solicitud_cursos.xlsx",
                            sheet = "A revisar por UA") %>% 
  select(c(1:19))
geografia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Geografía_Solicitud_cursos.xlsx",
                               sheet = "A revisar por UA",col_names = col_var)
historia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Historia_Solicitud_cursos.xlsx",
                              sheet = "A revisar por UA",col_names = col_var)
ingenieria_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Ingeniería_Solicitud_cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
letras_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Letras_Solicitud_cursos.xlsx",
                            sheet = "A revisar por UA",col_names = col_var)
matematicas_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Matemáticas_Solicitud_cursos.xlsx",
                                 sheet = "A revisar por UA",col_names = col_var)
#medicina_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Medicina_Solicitud_cursos.xlsx",
 #                             sheet = "A revisar por UA",col_names = col_var)
musica_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Música_Solicitud_cursos.xlsx",
                            sheet = "A revisar por UA",col_names = col_var)
psicologia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Psicología_Solicitud_cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
quimica_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Química_Solicitud_cursos.xlsx",
                             sheet = "A revisar por UA",col_names = col_var)
sociologia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Sociología_Solicitud_cursos.xlsx",
                                sheet = "A revisar por UA",col_names = col_var)
teologia_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Teología_Solicitud_cursos.xlsx",
                              sheet = "A revisar por UA",col_names = col_var)
trabajo_social_rev_UA <- read_excel("bd_2024-1/UAS_14-08/Trabajo Social_Solicitud_cursos.xlsx",
                                    sheet = "A revisar UA")

actuacion_sobre <- read_excel("bd_2024-1/UAS_14-08/Actuación_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG")

col_var2 <- colnames(actuacion_sobre)
agronomia_sobre <- read_excel("bd_2024-1/UAS_14-08/Agronomía_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
antropologia_sobre <- read_excel("bd_2024-1/UAS_14-08/Antropología_Solicitud cursos.xlsx",
                                 sheet = "Sobrepasos DMG",col_names = col_var2)
arquitectura_sobre <- read_excel("bd_2024-1/UAS_14-08/Arquitectura_Solicitud cursos.xlsx",
                                 sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
arte_sobre <- read_excel("bd_2024-1/UAS_14-08/Arte_Solicitud cursos.xlsx",
                         sheet = "Sobrepasos DMG",col_names = col_var2)
astrofisica_sobre <- read_excel("bd_2024-1/UAS_14-08/Astrofísica_Solicitud cursos.xlsx",
                                sheet = "Sobrepasos DMG",col_names = col_var2)
ciencia_politica_sobre <- read_excel("bd_2024-1/UAS_14-08/Ciencia Política_Solicitud cursos.xlsx",
                                     sheet = "Sobrepasos DMG",col_names = col_var2)
ciencias_biologicas_sobre <- read_excel("bd_2024-1/UAS_14-08/Ciencias Biológicas_Solicitud cursos.xlsx",
                                        sheet = "Sobrepasos DMG",col_names = col_var2)
ciencias_salud_sobre <- read_excel("bd_2024-1/UAS_14-08/Ciencias de la Salud_Solicitud cursos.xlsx",
                                   sheet = "Sobrepasos DMG",col_names = col_var2)
comunicaciones_sobre <- read_excel("bd_2024-1/UAS_14-08/Comunicaciones_Solicitud cursos.xlsx",
                                   sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
construccion_civil_sobre <- read_excel("bd_2024-1/UAS_14-08/Construcción Civil_Solicitud cursos.xlsx",
                                       sheet = "Sobrepasos DMG",col_names = col_var2)
derecho_sobre <- read_excel("bd_2024-1/UAS_14-08/Derecho_Solicitud cursos.xlsx",
                            sheet = "Sobrepasos DMG",col_names = col_var2)
des_sustentable_sobre <- read_excel("bd_2024-1/UAS_14-08/Desarrollo Sustentable_Solicitud cursos.xlsx",
                                    sheet = "Sobrepasos DMG",col_names = col_var2)
diseño_sobre <- read_excel("bd_2024-1/UAS_14-08/Diseño_Solicitud cursos.xlsx",
                           sheet = "Sobrepasos DMG",col_names = col_var2)
eco_admin_sobre <- read_excel("bd_2024-1/UAS_14-08/Economía y Administración_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
educacion_sobre <- read_excel("bd_2024-1/UAS_14-08/Educación_Solicitud cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
enfermeria_sobre <- read_excel("bd_2024-1/UAS_14-08/Enfermería_Solicitud cursos.xlsx",
                               sheet = "Sobrepasos DMG",col_names = col_var2)
esc_gobierno_sobre <- read_excel("bd_2024-1/UAS_14-08/Escuela de Gobierno_Solicitud_cursos.xlsx",
                                 sheet = "Sobrepasos DMG",col_names = col_var2)
estetica_sobre <- read_excel("bd_2024-1/UAS_14-08/Estética_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
est_urb_sobre <- read_excel("bd_2024-1/UAS_14-08/Estudios Urbanos_Solicitud_cursos.xlsx",
                            sheet = "Sobrepasos DMG",col_names = col_var2)
filosofia_sobre <- read_excel("bd_2024-1/UAS_14-08/Filosofía_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
fisica_sobre <- read_excel("bd_2024-1/UAS_14-08/Física_Solicitud_cursos.xlsx",
                           sheet = "Sobrepasos DMG",col_names = col_var2)
geografia_sobre <- read_excel("bd_2024-1/UAS_14-08/Geografía_Solicitud_cursos.xlsx",
                              sheet = "Sobrepasos DMG",col_names = col_var2)
historia_sobre <- read_excel("bd_2024-1/UAS_14-08/Historia_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
ingenieria_sobre <- read_excel("bd_2024-1/UAS_14-08/Ingeniería_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
letras_sobre <- read_excel("bd_2024-1/UAS_14-08/Letras_Solicitud_cursos.xlsx",
                           sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
matematicas_sobre <- read_excel("bd_2024-1/UAS_14-08/Matemáticas_Solicitud_cursos.xlsx",
                                sheet = "Sobrepasos DMG",col_names = col_var2)
medicina_sobre <- read_excel("bd_2024-1/UAS_14-08/Medicina_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
musica_sobre <- read_excel("bd_2024-1/UAS_14-08/Música_Solicitud_cursos.xlsx",
                           sheet = "Sobrepasos DMG") %>% 
  select(c(1:17))
psicologia_sobre <- read_excel("bd_2024-1/UAS_14-08/Psicología_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG",col_names = col_var2)
quimica_sobre <- read_excel("bd_2024-1/UAS_14-08/Química_Solicitud_cursos.xlsx",
                            sheet = "Sobrepasos DMG",col_names = col_var2)
sociologia_sobre <- read_excel("bd_2024-1/UAS_14-08/Sociología_Solicitud_cursos.xlsx",
                               sheet = "Sobrepasos DMG",col_names = col_var2)
teologia_sobre <- read_excel("bd_2024-1/UAS_14-08/Teología_Solicitud_cursos.xlsx",
                             sheet = "Sobrepasos DMG",col_names = col_var2)
trabajo_social_sobre <- read_excel("bd_2024-1/UAS_14-08/Trabajo Social_Solicitud_cursos.xlsx",
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
                            #medicina_rev_UA,
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
                           #medicina_sobre,
                           musica_sobre,
                           psicologia_sobre,
                           quimica_sobre,
                           sociologia_sobre,
                           teologia_sobre,
                           trabajo_social_sobre
) %>% 
  arrange(`Nombre estudiante`)

consolidado_sobre2 <- consolidado_sobre %>% 
  mutate("Respuesta UA" = rep("NO APLICA",2195)) %>% 
  mutate("Observación UA" = rep("NO APLICA",2195)) %>% 
  select(c(1:15),c(18:19),c(17,16)) %>% 
  rename("Fecha de solicitud del estudiante" = `Fecha de solicitud`)

consolidado_rev_UA <- consolidado_rev_UA %>% 
  filter(`Nombre estudiante` != "Nombre completo" & `Nombre estudiante` != "NA" &
           `Nombre estudiante` != "Nombre estudiante")%>% 
  arrange(`Nombre estudiante`)

# seleccionamos el nombre de cada estudiante de forma única
estudiantes <- unique(consolidado_rev_UA$`Nombre estudiante`)

# estos es transversal para todos los correos
source("codigos/codigo_solicitar_contraseña.R")
usuario <- "mbsaavedra@uc.cl"
contrasenya <- get_password()
asunto <- "[IMPORTANTE] Actualización solicitudes de cursos UC"

# comenzamos a escribir cada correo

for (i in estudiantes) {
  # filtramos el consolidado para el estudiante i
  cursos_aux <- filter(consolidado_rev_UA,`Nombre estudiante` == i)
  
  # creamos una tabla con la info de los cursos para el estudiante i
  if (nrow(cursos_aux) != 0) {
    # se le da el correo del estudiante i 
    para <- cursos_aux_UA$`Correo electrónico`[1]
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
  
  cursos_aux_UA <- filter(consolidado_rev_UA,`Nombre estudiante` == i)
  
  if (nrow(cursos_aux_UA) != 0) {
    # se le da el correo del estudiante i 
    para <- cursos_aux_UA$`Correo electrónico`[1]
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
   
<p>Junto con saludar y esperando que se encuentren muy bien, les enviamos a continuación la actualización de los cursos solicitados:</p>",
tabla_rev_UA,
"<h3><u>Importante:</u></h3>
<p>Los estados de la solicitud de cada curso significan:</p>
  <ul>
  <li><strong>Sin revisar:</strong> hasta el momento la Unidad Académica no ha revisado la solicitud. Si aún estás interesado/a en inscribir este curso, responde el formulario que enviaremos dentro de las próximas horas.</li>
  <li><strong>En revisión:</strong> la Unidad Académica está revisando la solicitud. Si aún estás interesado/a en inscribir este curso, responde el formulario que enviaremos dentro de las próximas horas.</li>
  <li><strong>Autorizado:</strong> la Unidad Académica sí autorizó una vacante adicional, por lo que podrás inscribir el curso. Si al intentar inscribir el curso en el sistema aparece “error”, se debe a que aún se está gestionando el reajuste de la vacante en el sistema. Nosotros gestionaremos internamente la inscripción de este curso en tu carga académica durante los próximos días. De todas maneras, debes comenzar a asistir a las clases según el horario disponible que puedes encontrar en <a href='https://buscacursos.uc.cl/'>buscacursos.uc.cl</a></li>
  <li><strong>Autorizado*:</strong>La Unidad Académica autorizó una vacante en una la sección distinta solicitada inicialmente.</li>
  <li><strong>No autorizado:</strong> la Unidad Académica no autorizó una vacante adicional, por lo que no podrás inscribir el curso. La no autorización de un curso puede deberse a distintas situaciones como falta de espacio en sala, requisitos previos no cumplidos, entre otros. </li>
  </ul>
<p>Que tengan una excelente tarde.  </p>",
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

