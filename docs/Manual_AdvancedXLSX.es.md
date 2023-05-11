



# Opciones avanzadas para XLSX
  
Da formato a celdas, crea y remueve hojas, filtra datos, agrega y elimina columnas y filas, abre archivos xls y transformalos a formato xlsx.  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Português](Manual_AdvancedXLSX.pr.md), [Español](Manual_AdvancedXLSX.es.md)*
  
![banner](imgs/Banner_AdvancedXLSX.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Descripción de los comandos

### Abrir xls
  
Abre un archivo xls para trabajar con el comando nativo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS |Selecciona el archivo XLS que quieres abrir|example.xls|
|Columna/as como fecha (opcional) ||0|
|Id (opcional) |Identificador de sesión|id|
|Asignar resultado a variable||Variable|

### Convertir xls a xlsx
  
Convierte un archivo formato xls a formato xlsx
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS |Selecciona el archivo XLS que quieres abrir|path/to/file/example.xls|
|Ruta archivo XLSX |Coloque la ruta completa donde quiere guardar el archivo XLSX (incluyendo nombre y extensión '.xlsx')|path/to/file/example.xlsx|

### Formatear celdas
  
Dar formato a celdas
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Sheet1|
|Celdas |Rango de celdas|A1:B5|
|Alineación Horizontal||---- Select ----|
|Alineación Vertical||---- Select ----|
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
