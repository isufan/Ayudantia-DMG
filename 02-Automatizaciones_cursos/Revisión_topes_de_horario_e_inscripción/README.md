# Automatización revisión topes de horario e inscripción
**Este código se encarga de seleccionar los dos primeros cursos con mayor prioridad indicado por cada estudiante, que no tengan tope y que se encuentren en el catálogo VRAI.** 

Para definir el orden en el cual se revisan los pares de cursos, se asignaron los siguientes "pesos" a cada prioridad (formato "prioridad : peso"): `{1 : 150, 2 : 130, 3 : 110, 4 : 100, 5 : 70, 6 : 60, 7 : 50, 8 : 40, 9 : 30, 10 : 20}`. Se puede ver que mientras mayor sea la prioridad, mayor será el peso. Se partirá revisando los cursos cuya suma de pesos sea mayor, hasta llegar a los cursos con menor prioridad.


Por ejemplo, primero se revisan los cursos con prioridad 1 y 2, luego 1 y 3, y, en caso de empate (como el 2 y 3, y 1 y 4), se revisa primero el que tiene el primer curso con mayor prioridad, es decir el 1 y 4. Los últimos cursos a revisar (en caso de ser necesario) serán aquellos con prioridad 9 y 10.

## Consideraciones importantes
- Este código se encarga de revisar que los dos cursos seleccionados no tengan tope en ningún módulo, sin importar si son cátedras, ayudantías, laboratorios, talleres u otros. También revisa que en caso de haber 2 módulos seguidos con actividades, no sean en campus distintos.
- Este código, a diferencia del código de la otra automatización, genera un archivo .xlsx de inmediato. También registra los casos con éxito y los sin éxito en hojas aparte. Los casos donde puede no haber éxito son cuando el estudiante tenga menos de dos opciones de cursos, o cuando no se encontraron dos cursos compatibles.
- Se realiza la verificación de que ambos cursos escogidos no tengan la misma sigla.
