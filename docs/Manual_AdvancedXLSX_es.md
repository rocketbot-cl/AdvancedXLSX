# Opciones avanzadas para XLSX
  
Módulo con opciones avanzadas para XLSX  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Español](Manual_AdvancedXLSX_es.md), [Portugues](Manual_AdvancedXLSX_pr.md).* 
  
![banner](imgs/Banner_AdvancedXLSX.png)

## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  



## Descripción de los comandos

### Abrir xls
  
Abre un archivo xls para trabajar con el comando nativo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS |Selecciona el archivo XLS que quieres abrir|example.xls|
|Columna/as como fecha (opcional) ||0|
|Id (opcional) |Identificador de sesión|id|
|Asignar resultado a variable||Variable|

### Formatear celdas
  
Dar formato a celdas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Sheet1|
|Celdas |Rango de celdas|A1:B5|
|Format ID |ID Formato. Ver Documentación https//learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1|0|
|Asignar resultado a variable||Variable|

### Crear hoja
  
Crea una nueva hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja que se creará|Sheet2|

### Borrar hoja
  
Borrar una hoja del libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja ||Sheet1|

### Contar en rango
  
Retorna el la máxima cantidad de filas y columnas desde una celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja donde se encuentra el rango|Sheet1|
|Celda de inicio|Celda de inicio del rango|A1|
|Asignar resultado a variable (Fila)|Nombre de variable donde se guardará el largo de la fila|Variable|
|Asignar resultado a variable (Columna)|Nombre de variable donde se guardará el largo de la columna|Variable|

### Filtrar por columna
  
Filtrar por columna
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Filtros |Filtros a aplicar.|["A > 3", "D *ARS", "C == Factura"]|
|Nombre de la hoja |Nombre de la hoja a filtrar.|Sheet1|
|Resultado detallado|Marcar para obtener resultado detallado.|True|
|Asignar resultado a variable||Variable|

### Eliminar Fila/Columna
  
Comando para eliminar filas y/o columnas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja a la que se le eliminará la fila o columna|Sheet1|
|Fila(s)|Rango de filas a eliminar|1:5|
|Columna(s)|Rango de columnas a eliminar|A:G|

### Insertar Fila/Columna
  
Comando para insertar filas y/o columnas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja a la que se le eliminará la fila o columna|Sheet1|
|Fila(s)|Rango de filas a eliminar|1:5|
|Columna(s)|Rango de columnas a eliminar|A:G|
