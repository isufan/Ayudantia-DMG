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

actuacion <- read_excel(
  "bd_2024-1/Actuación_Solicitud cursos.xlsx",
  sheet = "A revisar por UA")

agronomia <- read_excel(
  "bd_2024-1/Agronomía_Solicitud cursos.xlsx",
  sheet = "A revisar por UA")

antropologia <- read_excel(
  "bd_2024-1/Antropología_Solicitud cursos.xlsx",
  sheet = "A revisar por UA")
# Despues se agregarian los códigos para los otros excels
arquitectura <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
arte <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                   sheet = "A revisar por UA")
astrofisica <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
ciencia_politica <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
ciencias_biologicas <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
ciencias_salud <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
comunicaciones <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
construccion_civil <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
deportes <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                       sheet = "A revisar por UA")
derecho <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                      sheet = "A revisar por UA")
des_sustentable <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
diseño <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                     sheet = "A revisar por UA")
eco_admin <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                        sheet = "A revisar por UA")
educacion <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                        sheet = "A revisar por UA")
enfermeria <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
esc_gobierno <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
estetica <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                       sheet = "A revisar por UA")
est_urb <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                      sheet = "A revisar por UA")
filosofia <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                        sheet = "A revisar por UA")
fisica <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                     sheet = "A revisar por UA")
geografia <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                        sheet = "A revisar por UA")
historia <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                       sheet = "A revisar por UA")
ing_mat_compu <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
ingenieria <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
ingenieria_bio_med <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
instituto_eticas_apli <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
letras <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                     sheet = "A revisar por UA")
matematicas <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
medicina <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                       sheet = "A revisar por UA")
medicina_vet <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
musica <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                     sheet = "A revisar por UA")
odontologia <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
psicologia <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
quimica <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                      sheet = "A revisar por UA")
sociologia <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")
teologia <- read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
                       sheet = "A revisar por UA")
trabajo_social <-
  read_excel("bd_2024-1/Actuación_Solicitud cursos.xlsx",
             sheet = "A revisar por UA")

# combinar todos los excel

consolidado <- rbind(actuacion,agronomia,antropologia
                     #,arquitectura,
                     #arte,
                     #astrofisica,
                     #ciencia_politica,
                     #ciencias_biologicas,
                     #ciencias_salud,
                     #comunicaciones,
                     #construccion_civil,
                     #deportes,
                     #derecho,
                     #des_sustentable,
                     #diseño,
                     #eco_admin,
                     #educacion,
                     #enfermeria,
                     #esc_gobierno,
                     #estetica,
                     #est_urb,
                     #filosofia,
                     #fisica,
                     #geografia,
                     #historia,
                     #ing_mat_compu,
                     #ingenieria,
                     #ingenieria_bio_med,
                     #instituto_eticas_apli,
                     #letras,
                     #matematicas,
                     #medicina,
                     #medicina_vet,
                     #musica,
                     #odontologia,
                     #psicologia,
                     #quimica,
                     #sociologia,
                     #teologia,
                     #trabajo_social
                     ) %>% 
  arrange(`Nombre completo`) %>% 
  # se agrega esta para probar el destinatario
  # despues se debería borrar
  mutate(correo = c(rep("ffvilcas@outlook.com", 72))) 

# seleccionamos el nombre de cada estudiante de forma única

estudiantes <- unique(consolidado$`Nombre completo`)

# estos es transversal para todos los correos

source("codigo_solicitar_contraseña.R")
usuario <- "fvilca@uc.cl"
contrasenya <- get_password()
asunto <- "[IMPORTANTE] Actualización solicitudes de cursos UC"

# comenzamos a escribir cada correo

for (i in estudiantes) {
  # filtramos el consolidado para el estudiante i
  cursos_aux <- filter(consolidado,`Nombre completo` == i )
  # se le da el correo del estudiante i 
  para <- cursos_aux$correo[1]
  
  # creamos una tabla con la info de los cursos para el estudiante i
  tabla <- cursos_aux %>% 
    select(`Nombre `,Sigla,NRC,`Respuesta UA`) %>%
    kable(align = "c",booktabs = TRUE,
          col.names = c("Nombre Curso","Sigla", "NRC",
                        "Respuesta Unidad Académica"))  %>%   
    row_spec(0, bold = T, color = "white", background = "#173F8A") %>%  
    kable_minimal()
  
  # se crea el cuerpo del correo en formato HTML
  
  mensaje <- str_c(
    "<p>Estimados y estimadas estudiantes,</p>",
    "<p>Junto con saludar y esperando que se encuentren muy bien, les enviamos a continuación la actualización de los cursos solicitados durante estos días.</p>",
    tabla,
    "<p><strong>Importante:</strong></p>",
    "<ul>",
    "<li>Si tienes un curso <strong>Autorizado</strong> debes inscribirlo a través de la plataforma de inscripción de cursos. Si la plataforma de inscripción no te permite inscribirlo, por favor, ponte en contacto con tu Coordinadora de Acompañamiento. De todas maneras, <strong>debes comenzar a asistir a las clases</strong> según el horario disponible que puedes encontrar en <a href='https://buscacursos.uc.cl/'>buscacursos.uc.cl</a></li>",
    "<li>Los <strong>No autorizados</strong> pueden responder a distintas situaciones: no se pueden abrir más cupos en el curso, no cuentas con el nivel apropiado para cursarlo, entre otros múltiples motivos.</li>",
    "<li>Para aquellos que tengan cursos <strong>Pendientes</strong> eso implica que la Unidad Académica aún no ha revisado tu caso, te agradecemos esperar al viernes 15 para recibir una actualización del estado. No debes venir a nuestra oficina nuevamente.</li>",
    "<li>Si hoy fuiste a la oficina de San Joaquín y tu solicitud no se ve reflejada arriba es porque este comunicado se programó antes de que vinieras, pero tu solicitud si está registrada.</li>",
    "<li>Si solicitaste cursos de Deportes te contactaremos directamente.</li>",
    "</ul>",
    "<p><strong>IMPORTANTE: Si la respuesta a tu curso llega hoy viernes 15 y no alcanzas a inscribir el curso por la plataforma, no te preocupes, te ayudaremos a inscribirlo la próxima semana. Por favor, revisar el correo que te enviamos hoy al respecto.</strong></p>",
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
