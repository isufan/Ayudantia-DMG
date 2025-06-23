# Automatizaciones cursos
Ambas automatizaciones se encuentran dentro de archivos .ipynb, que consisten en **Jupyter Notebooks**. Un Jupyter Notebook consiste en una aplicación web que permite dividir el código en diferentes bloques y trabajar en cada bloque por separado sin importar el orden. Luego, cada bloque también se puede ejecutar de manera independiente. Cada archivo .ipynb puede ser denominado como "cuaderno".
## ¿Como abrir y ejecutar el código?
En caso de no tener Jupyter Notebook instalado de forma local, los cuadernos también se pueden abrir y ejecutar de forma online. Para esto, los pasos son:
1. Abrir Google Colab (https://colab.research.google.com/) e iniciar sesión de Google si es necesario.
2. Al llegar a la ventana "Open notebook", ir a "Upload" y pinchar "Browse".
3. Seleccionar el archivo .ipynb deseado.
4. Con el archivo abierto, pinchar "Connect" (abajo del botón "Comment" en la esquina superior derecha). Esperar a que se conecte (toma unos pocos segundos, aparecerá un check verde una vez conectado).
5. Una vez conectado, ahora se deben importar en Colab los archivos que contienen la información del formulario y la información de los cursos. Para esto, pinchar el ícono de carpeta al extremo izquierdo de la pantalla y luego pinchar el ícono de la hoja con una flecha hacia arriba. Ahí se deben seleccionar las siguientes planillas excel:
    - En caso de querer ejecutar la **automatización de revisión de información de cursos:**
      - Catálogo cognos
      - Consolidado de preselección (que contiene las respuestas del formulario y el catálogo VRAI).
    - En caso de querer ejecutar la **automatización de revisión de topes de horario y selección de cursos preliminar:**
      - Consolidado de preselección (que contiene las respuestas del formulario y el catálogo VRAI).
8. Está todo listo para ejecutar los códigos. Para esto, simplemente se pueden pinchar los botones de play que aparecen al poner el mouse sobre cada bloque de código. Para este caso, los bloques de código se deben ejecutar de arriba hacia abajo.
9. Para ambas automatizaciones, al ejecutar el código se generará un archivo con los resultados. Este archivo aparecerá en el mismo lugar donde se importaron las planillas excel en el paso 5 (ícono de carpeta). Al poner el mouse sobre este archivo y pinchar los tres puntos que aparecen, se puede ver que está la opción de descargarlo.
10. ¡Eso es todo! :tada:

### Consideraciones importantes
- Las automatizaciones pueden tomar algo de tiempo en ejecutarse, especialmente al importar los datos desde las planillas, pero el tiempo no debería sobrepasar los 20 segundos.
- Poner especial atención a que los nombres de las planillas, hojas de las planillas y de columnas definidos en el código sean los mismos que los reales.
  - **En caso de que los nombres de los archivos difieran**, se puede definir el nombre correcto en el código en las variables `df_courses_global_file` y `df_courses_vrai_file`.
  - **En caso de que los nombres de las hojas de las planillas difieran**, se pueden definir los nombres correctos en el parámetro `sheet_name` al momento de llamar la función `read_excel`.
  - **En caso de que los nombres de las columnas difieran**, se pueden definir los nombres correctos en las variables en las variables `cols_courses_global`, `cols_courses_vrai`, `cols_students` y `cols_students_dpt` en la automatización de revisión de información de cursos, y en las variables `cols_df_students` y `cols_df_courses` en la automatización de revisión de topes de horario.
- A pesar de los esfuerzos de abarcar todos (o casi todos) los casos posibles, puede existir la posibilidad de que el código falle para ciertos casos muy específicos (por ejemplo celdas vacías/nulas, espacios de más). También ocurrió que la información del catálogo cognos era diferente a la del catálogo VRAI. Por esto, se recomienda hacer una revisión general de correctitud para asegurarse que la información que entregan los códigos sea la correcta.

#### En caso de cualquier consulta o inquietud, contactar a Ignacio Sufán - ignacio.sufan@uc.cl
