



# Opciones avanzadas para XLSX
  
Da formato a celdas, crea y remueve hojas, filtra datos, agrega y elimina columnas y filas, abre archivos xls y transformalos a formato xlsx.  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Português](Manual_AdvancedXLSX.pr.md), [Español](Manual_AdvancedXLSX.es.md)*
  
![banner](imgs/Banner_AdvancedXLSX.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  



## Como usar este módulo

Solo si utiliza la version 2023 de Rocketbot debe seguir los siguientes pasos para evitar el error:

ImportError: cannot import name 'etree' from 'lxml'

1. Debe dirigirse a la carpeta raiz de Rocketbot y validar que exista la libreria 'lxml'.
2. En caso que no exista, desde una terminal ir a la carpeta raiz de Rocketbot y colocar: 
pip install lxml -t .
3. Tomar en cuenta que, debe instalar la libreria con Python 3.10 de 64bits.


## Descripción de los comandos

### Abrir xls
  
Abre un archivo xls para trabajar con el comando nativo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS |Selecciona el archivo XLS que quieres abrir|example.xls|
|Columna/as como fecha (opcional) ||0|
|Id (opcional) |Identificador de sesión|id|
|Encoding|Tipo de Encoding a aplicar. Por defecto latin-1|latin-1|
|Asignar resultado a variable||Variable|

### Abrir xlsx avanzado
  
Abre un archivo xlsx para trabajar con el comando nativo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLSX |Selecciona el archivo XLSX que quieres abrir|example.xlsx|
|Solo lectura|Marque si desea abrir el xlsx solo para lectura, el contenido no se podrá editar.|False|
|Conservar vba|Marcar para conservar el posible codigo VBA que pudiera tener el libro.|False|
|Solo data|Controla si las celdas con fórmulas tienen la fórmula (predeterminado) o el valor almacenado la última vez que Excel leyó la hoja.|False|
|Conservar links|Marcar si se deben conservar los enlaces a libros de trabajo externos.|False|
|Id (opcional) |Identificador de sesión|id|
|Asignar resultado a variable||Variable|

### Convertir xls a xlsx
  
Convierte un archivo formato xls a formato xlsx
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo XLS |Selecciona el archivo XLS que quieres abrir|path/to/file/example.xls|
|Ruta archivo XLSX |Coloque la ruta completa donde quiere guardar el archivo XLSX (incluyendo nombre y extensión '.xlsx')|path/to/file/example.xlsx|
|Encoding|Tipo de Encoding a aplicar. Por defecto latin-1|latin-1|

### Convertir hoja a csv
  
Convierte una hoja del archivo xlsx abierto a csv
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo CSV |Selecciona el archivo CSV que quieres abrir|path/to/file/example.csv|
|Delimitador|Separador del archivo csv|,|
|Formato de salida de Fechas|Formato con el que se van a convertir las fechas de la Hoja xlsx a csv|%d/%m/%Y|
|Asignar resultado a variable|Nombre de la variable donde guardar el resultado|Variable|

### Leer rango
  
Devuelve el valor del rango dado. Un valor si el rango es una celda o una lista si el rango tiene varias celdas.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja |Nombre de la hoja donde se encuentra el rango|Sheet1|
|Celda o Rango|Celda de inicio del rango|A1|
|Asignar resultado a variable (Columna)|Nombre de variable donde se guardará el largo de la columna|Variable|

### Renombrar hoja
  
Renombrar una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la hoja a renombrar |Nombre que tiene la hoja a renombrar|OldSheet|
|Nuevo nombre de la hoja |Nombre que tendrá la hoja|NewSheet|

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
|Filtros |Filtros a aplicar. Para filtrar por vacíos indicar == None|["A > 3", "D *ARS", "C == Factura"]|
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

### Insertar imagen
  
Insertar una imagen en un documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta imagen |Selecciona el archivo de la imagen que quieres insertar en el documento|example.png|
|Hoja |Nombre de la hoja del documento donde insertar la imagen|Sheet1|
|Celda |Celda donde insertar la imagen|A1|

### Cerrar xlsx
  
Cerrar un archivo xlsx abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Asignar resultado a variable||Variable|
