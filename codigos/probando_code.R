# probando codigos

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

source("codigo_solicitar_contraseña.R")
usuario <- "fvilca@uc.cl"
contrasenya <- get_password()
para <- c("ffvilcas@outlook.com") #Podría ser una lista
asunto <- "Esta es una prueba hotmail"
mensaje <- "Esto es el cuerpo del mensaje de prueba desde hotmail"

smtp <- server(host = "smtp.office365.com",
               port = 587,
               username = usuario,
               password = contrasenya)

email <- envelope(
  from = usuario,
  to= para,
  subject = asunto,
  text = mensaje)

smtp(email, verbose = TRUE)

# REVISAR INFORMACION DE CURSOS

# NCR - Sigla - Nombre del form debe coincidir con el de cognos
# modificar si es necesario
# comprobar que la unidad académica corresponde

# cargar datos

cursos <- import("bd_2024-1/Cursos programados 2024_FINAL.xlsx")
View(cursos)

cursos_int <- import("bd_2024-1/Catálogo de cursos estudiantes intercambio 2024-1.xlsx")
View(cursos_int)

formulario <- read_excel("resultados/formulario1.xlsx")
View(formulario)

table(formulario$ID) # 199 estudiantes + 2 estudiantes q no tienen ID

cursos2 <- cursos %>% 
  select(NRC,Materia,`Número Curso`,`Nombre Curso`,Escuela)  %>%
  mutate(nrc = NRC) %>% 
  mutate(sigla = str_c(Materia,`Número Curso`)) %>% 
  mutate(escuela = str_sub(cursos$Escuela,6,-1)) %>% 
  select(nrc,sigla,Materia,`Número Curso`,`Nombre Curso`,escuela)

for (i in c(1:(nrow(cursos2)))) {
  cursos2$sigla[i] <- ifelse(nchar(cursos2$`Número Curso`[i]) ==1,
                     str_c(cursos2$Materia[i],"00",cursos2$`Número Curso`[i]),
                     ifelse(nchar(cursos2$`Número Curso`[i])==2,
                            str_c(cursos2$Materia[i],"0",cursos2$`Número Curso`[i]),
                            str_c(cursos2$Materia[i],cursos2$`Número Curso`[i])))
}


View(cursos2)

formulario_corr <- formulario
problemas <- 0
casos_revisar <- data.frame(
  Nombre_completo = character(),
  Nombre = character(),
  NRC = character(),
  Sigla = character(),
  Unidad_Academica = character())

for (i in seq_len(nrow(formulario))) {
  if (str_trim(formulario$NRC[i]) %in% cursos2$nrc) {
    if (str_trim(str_to_upper(formulario$Sigla[i])) %in% cursos2$sigla) {
      
      # Si entramos aquí significa que inscribió bien sigla y nrc
      # Falta corregir nombre y Unidad académica
      
      aux <- filter(cursos2, nrc == formulario$NRC[1] &
                      sigla == str_to_upper(formulario$Sigla[1]))
      
      if (nrow(aux) > 0) {
        formulario_corr$Nombre[i] <- aux[, 3]
        formulario_corr$`Unidad Académica`[i] <- aux[, 4]
      } 
    } else {
      # Si entramos aquí significa que inscribió bien NRC y mal la sigla
      # Se inscribe el curso según el NRC si está
      
      mensaje <- str_c("Estudiante ", formulario$`Nombre completo`[i],
                       " no coincide en la sigla ", formulario$Sigla[i])
      
      # Igual sumamos el problema para revisarlo
      print(mensaje)
      problemas <- problemas + 1
      
      aux <- filter(cursos2, nrc == formulario$NRC[i])
      casos_revisar <- rbind(casos_revisar,
                             c(formulario$`Nombre completo`[i],
                               formulario$`Nombre `[i],
                               formulario$NRC[i],
                               formulario$Sigla[i],
                               formulario$`Unidad Académica`[i]))
      
      if (nrow(aux) > 0) {
        formulario_corr$Sigla[i] <- aux[, 2]
        formulario_corr$Nombre[i] <- aux[, 3]
        formulario_corr$`Unidad Académica`[i] <- aux[, 4]
      }
    }
  } else if (str_trim(formulario$NRC[i]) == "Sin Info" |
             str_trim(formulario$NRC[i]) == "NA") {
    
    # Si entramos aquí significa que no inscribió NRC 
    mensaje <- str_c("Estudiante ", formulario$`Nombre completo`[i],
                     " no indica NRC")
    print(mensaje)
    problemas = problemas + 1
    casos_revisar <- rbind(casos_revisar,
                           c(formulario$`Nombre completo`[i],
                             formulario$`Nombre `[i],
                             formulario$NRC[i],
                             formulario$Sigla[i],
                             formulario$`Unidad Académica`[i]))
    
  } else if (str_trim(str_to_upper(formulario$Sigla[i])) %in% cursos2$sigla) {
    
    # Si entramos aquí significa que inscribió bien la sigla y mal el NRC
    # Se inscribe el curso según la sigla si está
    
    mensaje <- str_c("Estudiante ", formulario$`Nombre completo`[i],
                     " no coincide en el NRC ", formulario$NRC[i])
    
    # Igual sumamos el problema para revisarlo
    print(mensaje)
    problemas <- problemas + 1
    
    aux <- filter(cursos2, sigla == str_to_upper(formulario$Sigla[i]))
    
    casos_revisar <- rbind(casos_revisar,
                           c(formulario$`Nombre completo`[i],
                             formulario$`Nombre `[i],
                             formulario$NRC[i],
                             formulario$Sigla[i],
                             formulario$`Unidad Académica`[i]))
    
    if (nrow(aux) > 0) {
      formulario_corr$NRC[i] <- aux[, 1]
      formulario_corr$Nombre[i] <- aux[, 3]
      formulario_corr$`Unidad Académica`[i] <- aux[, 4]
    }
    
  } else {
    # Si entramos aquí significa que no coincide el NRC ni la sigla con
    # las del portal
    
    mensaje <- str_c("No se hay coincidencia exacta para el NRC ",
                     formulario$NRC[i], " y la sigla ",
                     formulario$Sigla[i], " para el estudiante ",
                     formulario$`Nombre completo`[i])
    print(mensaje)
    problemas <- problemas + 1
    
    casos_revisar <- rbind(casos_revisar,
                             c(formulario$`Nombre completo`[i],
                             formulario$`Nombre `[i],
                             formulario$NRC[i],
                             formulario$Sigla[i],
                             formulario$`Unidad Académica`[i]))
  }
}

print(str_c("hay ", problemas, " problemas"))
names(casos_revisar) <- c("Nombre_completo",
                          "Nombre","NRC","Sigla","Unidad_Academica")
                          
View(casos_revisar)

# Se guardan los excels obtenidos

# el excel con las variables esenciales de cognos
write_xlsx(cursos2, "resultados/cursos_cognos.xlsx")

# el excel con el formulario corregido
write_xlsx(formulario_corr, "resultados/formulario_corregido.xlsx")

# el excel con los casos a revisar con algun problema 
# en sigla o nrc 
write_xlsx(casos_revisar, "resultados/casos_a_revisar.xlsx")




