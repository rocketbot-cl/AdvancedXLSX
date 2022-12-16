# Opciones avanzadas para XLSX
  
Módulo con opciones avanzadas para XLSX  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Español](Manual_AdvancedXLSX.es.md).*
  
![banner](imgs/Banner_advancedxlsx.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  



## Descripción de los comandos

### Abrir xls
  
Abre un archivo xls para trabajar con el comando nativo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS|Selecciona el archivo XLS que quieres abrir|Archivo.XLS|
|Identificador (opcional)|Identificador de sesión|id|

### Obtener celda con formato
  
Obtiene una rango de celdas con el formato
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese celdas |Rango de celdas|A1:B5|
|Asignar resultado a variable|Nombre de variable donde se guardara el resultado|Variable|

### Crear hoja
  
Crea una nueva hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja que se creará|Hoja2|

### Contar en rango
  
Retorna el la máxima cantidad de filas y columnas desde una celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja donde se encuentra el rango|Hoja2|
|Celda de inicio|Celda de inicio del rango|A1|
|Asignar largo de fila a variable|Nombre de variable donde se guardará el largo de la fila|Variable|
|Asignar largo de columna a variable|Nombre de variable donde se guardará el largo de la columna|Variable|

### Filtrar por columna
  
Filtrar por columna
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Filtros |Filtros a aplicar.|["A > 3", "D *ARS", "C == Factura"]|
|Nombre de la hoja |Nombre de la hoja a filtrar.|hoja1|
|Resultado detallado|Marcar para obtener resultado detallado.|True|
|Variable donde almacenar resultado |Variable se almacenará el resultado.|resultado|

### Eliminar Fila/Columna
  
Comando para eliminar filas y/o columnas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja a la que se le eliminará la fila o columna|Hoja2|
|Fila(s)|Rango de filas a eliminar|1:5|
|Columna(s)|Rango de columnas a eliminar|A:G|
