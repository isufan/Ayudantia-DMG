# Codigos originales 
# 
# Proceso 2024-1

# Cargar los datos desde el archivo CSV
datos <- data
# Crear una nueva tabla con el formato deseado para cada estudiante
nueva_tabla <- data.frame()
# Recorrer cada fila de la tabla original
for (i in 1:nrow(datos)) {
# Obtener la información del estudiante actual
estudiante <- datos[i, ]
# Obtener las columnas relacionadas con los cursos (asumiendo que van desde C1 hasta C10)
cursos_cols <- grep("^C[0-9]+UA$", names(estudiante), value = TRUE)
# Iterar sobre las columnas de los cursos y agregar filas a la nueva tabla
for (col in cursos_cols) {
  # Filtrar las columnas de interés para el curso actual
curso_info <- estudiante[c("ID", "Hora de inicio", "Hora de finalización",
                           "Correo electrónico", "Nombre completo",
                           "Transcript of Record", "RUT", "DV", "RUT UC",
                           "Correo electrónico2", "Paisdeorigen",
                           "Universidad de Origen", "creditosuc", col, 
                           paste0(col, "N"),paste0(col, "NRC"), 
                           paste0(col, "Sigla"), paste0(col, "NS"), 
                           paste0(col, "GRADO"))]
  # Renombrar las columnas para que coincidan con el formato deseado
colnames(curso_info) <- c("ID", "Hora de inicio", "Hora de finalización",
                          "Correo electrónico", "Nombre completo",
                          "Transcript of Record", "RUT", "DV", "RUT UC",
                          "Correo electrónico2", "Paisdeorigen",
                          "Universidad de Origen", "creditosuc", "CUA",
                          "CN", "CNRC","CSigla", "CNS", "CGRADO")
  # Agregar la información del curso actual a la nueva tabla
  nueva_tabla <- rbind(nueva_tabla, curso_info)
}
}

# Cargar los datos desde el archivo CSV
datos <- read_excel("bd_2024-1/1. Formulario pre-selección de cursos intercambio 2024-1(1-195) (1).xlsx")
# Crear una nueva tabla con el formato deseado para cada estudiante
nueva_tabla <- data.frame()
# Recorrer cada fila de la tabla original
for (i in 1:nrow(datos)) {
  # Obtener la información del estudiante actual
  estudiante <- datos[i, ]
  
  # Obtener las columnas relacionadas con los cursos (asumiendo que van desde C1 hasta C10)
cursos_cols <- grep("^C[0-3]+UA$", names(estudiante), value = TRUE)

# Iterar sobre las columnas de los cursos y agregar filas a la nueva tabla
for (col in cursos_cols) {
  # Filtrar las columnas de interés para el curso actual
  curso_info <- estudiante[c("ID", "Hora de inicio", "Hora de finalización",
                             "Correo electrónico", "Nombre completo",
                             "Transcript of Record", "RUT", "DV", "RUT UC",
                             "Correo electrónico2","Paisdeorigen", 
                             "Universidad de Origen","creditosuc",
                             "TGICD","CCDTGI","selecciondeportiva",
                             "deporte","pruebasselecciondeportiva",
                             "Comentarios", col, paste0(col, "N"),
                             paste0(col, "NRC"),paste0(col, "Sigla"))]
  
  # Renombrar las columnas para que coincidan con el formato deseado
  colnames(curso_info) <- c("ID", "Hora de inicio", "Hora de finalización",
                            "Correo electrónico", "Nombre completo",
                            "Transcript of Record", "RUT", "DV", "RUT UC",
                            "Correo electrónico2", "Paisdeorigen", 
                            "Universidad de Origen", "creditosuc","TGICD",
                            "CCDTGI","selecciondeportiva","deporte",
                            "pruebasselecciondeportiva","Comentarios", "CUA",
                            "CN", "CNRC", "CSigla")
  
  # Agregar la información del curso actual a la nueva tabla
  nueva_tabla <- rbind(nueva_tabla, curso_info)
}
}

datos <- datos %>% group_by(nombre_completo) %>% mutate(numero_curso =
                                                          row_number())
