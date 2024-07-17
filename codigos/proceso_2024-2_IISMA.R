# Estudiantes de Cursos IISMA


library(readxl)
library(dplyr)
library(writexl)

# Cursos regulares ------

# Cargar los datos desde el archivo xlsx
datos <- read_excel("bd_2024-1/2. Formulario preselección de cursos IISMA.xlsx", sheet = "Respuestas formulario")


# Recodificamos las variables para desagregar cursos
# Estas pueden variar si es que el nombre de 
# las variables se modfica

cursos_inscritos <- datos %>% 
  rename(C0UA =`Course 1. Academic Unit`) %>% 
  rename(C1UA =`Course 2. Academic Unit`) %>%
  rename(C2UA =`Course 3. Academic Unit`) %>%
  rename(C3UA =`Course 4. Academic Unit`) %>%
  rename(C4UA =`Course 5. Academic Unit`) %>%
  rename(C5UA =`Course 6. Academic Unit`) %>%
  rename(C6UA =`Course 7. Academic Unit`) %>%
  rename(C0UAN = `Course 1. Name`) %>% 
  rename(C1UAN = `Course 2. Name`)  %>%
  rename(C2UAN = `Course 3. Name`)  %>%
  rename(C3UAN = `Course 4. Name`) %>%
  rename(C4UAN = `Course 5. Name`)  %>%
  rename(C5UAN = `Course 6. Name`) %>%
  rename(C6UAN = `Course 7. Name`) %>%
  rename(C0UANRC = `Course 1. NCR`) %>% 
  rename(C1UANRC = `Course 2. NCR`) %>%
  rename(C2UANRC = `Course 3. NCR`) %>%
  rename(C3UANRC = `Course 4. NCR`) %>%
  rename(C4UANRC = `Course 5. NCR`) %>%
  rename(C5UANRC = `Course 6. NCR`) %>%
  rename(C6UANRC = `Course 7. NCR`) %>%
  rename(C0UASIGLA =`Course 1. Code`) %>% 
  rename(C1UASIGLA =`Course 2. Code`) %>%
  rename(C2UASIGLA =`Course 3. Code`) %>%
  rename(C3UASIGLA =`Course 4. Code`) %>%
  rename(C4UASIGLA =`Course 5. Code`) %>%
  rename(C5UASIGLA =`Course 6. Code`) %>%
  rename(C6UASIGLA =`Course 7. Code`) %>%
  rename(C0UANS =`Course 1. Section number`) %>% 
  rename(C1UANS =`Course 2. Section number`) %>%
  rename(C2UANS =`Course 3. Section number`) %>%
  rename(C3UANS =`Course 4. Section number`) %>%
  rename(C4UANS =`Course 5. Section number`) %>%
  rename(C5UANS =`Course 6. Section number`) %>%
  rename(C6UANS =`Course 7. Section number`) %>% 
  rename(comentarios = `Comments `)


cursos_desagregados <- data.frame()

# Recorrer cada fila de la tabla original

for (i in 1:nrow(cursos_inscritos)) {
  
  # Obtener la información del estudiante actual
  
  estudiante <- cursos_inscritos[i, ]
  
  # Obtener las columnas relacionadas con los cursos 
  # (asumiendo que van desde C0 hasta C09)
  
  cursos_cols <- grep("^C[0-6]+UA$", names(estudiante), value = TRUE)
  # Iterar sobre las columnas de los cursos y 
  # agregar filas a la nueva tabla
  
  for (col in cursos_cols) {
    
    # Filtrar las columnas de interés para el curso actual
    
    curso_info <- estudiante[c("Fecha de solicitud",
                               "Nombre completo",
                               "Transcript of Record",
                               "RUT UC",
                               "Correo electrónico",
                               col, 
                               paste0(col, "N"),
                               paste0(col, "NRC"), 
                               paste0(col, "SIGLA"),
                               paste0(col, "NS"),
                               "comentarios")]
    
    # Renombrar las columnas para que coincidan 
    # con el formato deseado
    
    colnames(curso_info) <- c("Fecha de solicitud",
                              "Nombre completo",
                              "Transcript of Record", "RUT UC",
                              "Correo electrónico",
                              "CUA","CN", "CNRC",
                              "CSigla", "CNS", "Comentarios")
    # Agregar la información del curso actual a la nueva tabla
    cursos_desagregados <- rbind(cursos_desagregados, curso_info)
  }
}

# Se renombran nuevamente las variables para que sean comprensibles
# para todo el mundo y para la siguiente etapa

cursos_desagregados <- cursos_desagregados %>% 
  rename("Unidad Académica" = CUA) %>%
  rename("Nombre" = CN) %>%
  rename("NRC" = CNRC) %>%
  rename("Sigla" = CSigla) %>%
  rename("Número de la sección" = CNS)

# Se eliminan los cursos que tienen todas las variables deseadas

info_vital <- c("Unidad Académica","Nombre", "NRC", "Sigla")

cursos_desagregados <- cursos_desagregados[!rowSums(is.na(cursos_desagregados[info_vital])) == length(info_vital), ]

# Enumerar cursos desagregados

cursos_enumerados <- cursos_desagregados %>%  
  group_by(`Nombre completo`) %>%  
  mutate(numero_curso = row_number())

# Se guarda la tabla en formato Excel

write_xlsx(cursos_enumerados, "resultados/formularioIISMA.xlsx")  
