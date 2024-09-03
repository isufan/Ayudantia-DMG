# Codigos originales 
# 
# Proceso 2024-2
#
# Para estudiantes rezagados de Agosto

library(readxl)
library(dplyr)
library(writexl)

# Cursos regulares ------

# Cargar los datos desde el archivo xlsx
datos <- read_excel("bd_2024-1/Formulario preselección de cursos intercambio 2024-2..(1-6).xlsx")

# Eliminamos las primeras variables y renombramos

datos <- datos %>% 
  mutate("Fecha de solicitud" = rep("Agosto",6)) %>% 
  select(c(91,7:90))

# Recodificamos las variables para desagregar cursos
# Estas pueden variar si es que el nombre de 
# las variables se modfica

cursos_inscritos <- datos %>% 
  rename(C0UA =`Curso 1. Unidad Académica`) %>% 
  rename(C1UA =`Curso 2. Unidad Académica`) %>%
  rename(C2UA =`Curso 3. Unidad Académica`) %>%
  rename(C3UA =`Curso 4. Unidad Académica`) %>%
  rename(C4UA =`Curso 5. Unidad Académica`) %>%
  rename(C5UA =`Curso 6. Unidad Académica`) %>%
  rename(C6UA =`Curso 7. Unidad Académica`) %>%
  rename(C7UA =`Curso 8. Unidad Académica`) %>%
  rename(C8UA =`Curso 9. Unidad Académica`) %>%
  rename(C9UA =`Curso 10. Unidad Académica`) %>% 
  rename(C0UAN =`Curso 1. Nombre `) %>% 
  rename(C1UAN =`Curso 2. Nombre `) %>%
  rename(C2UAN =`Curso 3. Nombre `) %>%
  rename(C3UAN =`Curso 4. Nombre`) %>%
  rename(C4UAN =`Curso 5. Nombre `) %>%
  rename(C5UAN =`Curso 6. Nombre`) %>%
  rename(C6UAN =`Curso 7. Nombre`) %>%
  rename(C7UAN =`Curso 8. Nombre`) %>%
  rename(C8UAN =`Curso 9. Nombre `) %>%
  rename(C9UAN =`Curso 10. Nombre`) %>% 
  rename(C0UANRC =`Curso 1. NRC`) %>% 
  rename(C1UANRC =`Curso 2. NRC`) %>%
  rename(C2UANRC =`Curso 3. NRC`) %>%
  rename(C3UANRC =`Curso 4. NRC `) %>%
  rename(C4UANRC =`Curso 5. NRC `) %>%
  rename(C5UANRC =`Curso 6. NRC`) %>%
  rename(C6UANRC =`Curso 7. NRC`) %>%
  rename(C7UANRC =`Curso 8. NRC`) %>%
  rename(C8UANRC =`Curso 9. NRC`) %>%
  rename(C9UANRC =`Curso 10. NRC`) %>% 
  rename(C0UASIGLA =`Curso 1. Sigla`) %>% 
  rename(C1UASIGLA =`Curso 2. Sigla`) %>%
  rename(C2UASIGLA =`Curso 3. Sigla`) %>%
  rename(C3UASIGLA =`Curso 4. Sigla `) %>%
  rename(C4UASIGLA =`Curso 5. Sigla `) %>%
  rename(C5UASIGLA =`Curso 6. Sigla`) %>%
  rename(C6UASIGLA =`Curso 7. Sigla`) %>%
  rename(C7UASIGLA =`Curso 8. Sigla`) %>%
  rename(C8UASIGLA =`Curso 9. Sigla`) %>%
  rename(C9UASIGLA =`Curso 10. Sigla`) %>% 
  rename(C0UANS =`Curso 1. Número de la sección`) %>% 
  rename(C1UANS =`Curso 2. Número de la sección`) %>%
  rename(C2UANS =`Curso 3. Número de la sección `) %>%
  rename(C3UANS =`Curso 4. Número de la sección `) %>%
  rename(C4UANS =`Curso 5. Número de la sección`) %>%
  rename(C5UANS =`Curso 6. Número de la sección`) %>%
  rename(C6UANS =`Curso 7. Número de la sección`) %>%
  rename(C7UANS =`Curso 8. Número de la sección`) %>%
  rename(C8UANS =`Curso 9. Número de la sección`) %>%
  rename(C9UANS =`Curso 10. Número de la sección`) %>% 
  rename(C0UAGRADO = `Curso 1. ¿Necesitas este curso para la obtención del grado en tu universidad de origen? `) %>% 
  rename(C1UAGRADO = `Curso 2. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C2UAGRADO = `Curso 3. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C3UAGRADO = `Curso 4. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C4UAGRADO = `Curso 5. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C5UAGRADO = `Curso 6. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C6UAGRADO = `Curso 7. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C7UAGRADO = `Curso 8. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C8UAGRADO = `Curso 9.¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>%
  rename(C9UAGRADO = `Curso 10. ¿Necesitas este curso para la obtención del grado en tu universidad de origen?`) %>% 
  rename("Total créditos UC" = `Cantidad de "créditos UC" totales que deseas/estimas cursar en la UC`)

# Crear una nueva tabla con el formato deseado para cada estudiante

cursos_desagregados <- data.frame()

# Recorrer cada fila de la tabla original

for (i in 1:nrow(cursos_inscritos)) {
  
  # Obtener la información del estudiante actual
  
  estudiante <- cursos_inscritos[i, ]
  
  # Obtener las columnas relacionadas con los cursos 
  # (asumiendo que van desde C0 hasta C09)
  
  cursos_cols <- grep("^C[0-9]+UA$", names(estudiante), value = TRUE)
  # Iterar sobre las columnas de los cursos y 
  # agregar filas a la nueva tabla
  
  for (col in cursos_cols) {
    
    # Filtrar las columnas de interés para el curso actual
    
    curso_info <- estudiante[c("Fecha de solicitud",
                               "Nombre completo",
                               "Correo electrónico2",
                               "País de tu universidad",
                               "Universidad de Origen",
                               "Total créditos UC",col, 
                               paste0(col, "N"),
                               paste0(col, "NRC"), 
                               paste0(col, "SIGLA"),
                               paste0(col, "NS"), 
                               paste0(col, "GRADO"))]
    
    # Renombrar las columnas para que coincidan 
    # con el formato deseado
    
    colnames(curso_info) <- c("Fecha de solicitud",
                              "Nombre estudiante",
                              "Correo electrónico",
                              "Pais de su universidad",
                              "Universidad de Origen",
                              "Total créditos UC", "CUA",
                              "CN", "CNRC","CSigla", "CNS", "CGRADO")
    # Agregar la información del curso actual a la nueva tabla
    cursos_desagregados <- rbind(cursos_desagregados, curso_info)
  }
}

# Se renombran nuevamente las variables para que sean comprensibles
# para todo el mundo y para la siguiente etapa

cursos_desagregados <- cursos_desagregados %>% 
  rename("Nombre completo" =`Nombre estudiante`) %>% 
  rename("Unidad Académica" = CUA) %>%
  rename("Nombre" = CN) %>%
  rename("NRC" = CNRC) %>%
  rename("Sigla" = CSigla) %>%
  rename("Número de la sección" = CNS) %>%
  rename("¿Necesitas este curso para la obtención del grado en tu universidad de origen?" = CGRADO)

# Se eliminan los cursos que tienen todas las variables deseadas

info_vital <- c("Unidad Académica","Nombre", "NRC", "Sigla")

cursos_desagregados <- cursos_desagregados[!rowSums(is.na(cursos_desagregados[info_vital])) == length(info_vital), ]

cursos_enumerados <- cursos_desagregados %>%  
  group_by(`Nombre completo`) %>%  
  mutate(numero_curso = row_number()) %>% 
  select(c(1:5),c(7:12),6)

# Se guarda la tabla en formato Excel

write_xlsx(cursos_enumerados, 
           "resultados/formulario_rezagados_agosto.xlsx")
