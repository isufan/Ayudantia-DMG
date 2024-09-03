# Cargar el paquete tcltk
library(tcltk)

# Función para solicitar la contraseña
get_password <- function(prompt = "Por favor, ingresa tu contraseña:") {
  # Crear ventana
  tt <- tktoplevel()
  tkwm.title(tt, "Ingresar Contraseña")
  
  # Crear variable de contraseña
  pw_var <- tclVar("")
  
  # Crear función para cerrar la ventana y obtener la contraseña
  on_ok <- function() {
    tkgrab.release(tt)
    tkdestroy(tt)
  }
  
  # Crear etiqueta y cuadro de entrada para la contraseña
  tkgrid(tklabel(tt, text = prompt), padx = 10, pady = 10)
  tkgrid(tkentry(tt, textvariable = pw_var, show = "*"), padx = 10, pady = 10)
  
  # Crear botón OK
  tkgrid(tkbutton(tt, text = "OK", command = on_ok), padx = 10, pady = 10)
  
  # Forzar ventana a primer plano
  tkgrab.set(tt)
  tkwait.window(tt)
  
  # Obtener valor de la contraseña
  password <- tclvalue(pw_var)
  
  # Limpiar variable de contraseña
  tclvalue(pw_var) <- ""
  
  return(password)
}
