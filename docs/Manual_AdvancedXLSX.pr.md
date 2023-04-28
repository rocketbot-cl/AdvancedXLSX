



# Opções avançadas para XLSX
  
Formate células, crie e remova planilhas, filtre dados, adicione e exclua colunas e linhas, abra arquivos xls e transforme-os no formato xlsx.  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Português](Manual_AdvancedXLSX.pr.md), [Español](Manual_AdvancedXLSX.es.md)*
  
![banner](imgs/Banner_AdvancedXLSX.png)
## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Descrição do comando

### Abrir xls
  
Abra um arquivo xls para trabalhar com o comando nativo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo XLS|Selecione o arquivo XLS que deseja abrir|example.xls|
|Coluna/as como data (opcional)||0|
|Id (optional) |Identificador de sessão|id|
|Atribuir resultado à variável||Variável|

### Converter xls para xlsx
  
Converter um arquivo de formato xls para o formato xlsx
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo XLS|Selecione o arquivo XLS que deseja abrir|path/to/file/example.xls|
|Caminho do arquivo XLSX|Coloque o caminho completo onde deseja salvar o arquivo XLSX (incluindo nome e extensão '.xlsx')|path/to/file/example.xlsx|

### Formatear celular
  
Dar formato a células
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Folha ||Sheet1|
|Células |Intervalo de células|A1:B5|
|Alinhamento Horizontal||---- Select ----|
|Alinhamento Vertical||---- Select ----|
|Format ID |Formato de ID. Ver documentação https//learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1|0|
|Atribuir resultado à variável||Variável|

### Criar folha
  
Criar uma nova folha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha |Nome da folha a ser criada|Sheet2|

### Excluir folha
  
Excluir uma folha de pasta de trabalho
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha||Sheet1|

### Contar no intervalo
  
Retorna o número máximo de linhas e colunas de uma célula
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha |Nome da planilha onde o intervalo está localizado|Sheet1|
|Célula inicial|Célula inicial de intervalo|A1|
|Atribuir resultado à variável (Linha)|Nome da variável onde o comprimento da linha será salvo|Variável|
|Atribuir resultado à variável (Coluna)|Nome da variável onde o comprimento da coluna será salvo|Variável|

### Filtrar por coluna
  
Filtrar por coluna
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Filtros|Filtros a aplicar.|["A > 3", "D *ARS", "C == Invoice"]|
|Nome da folha|Nome da folha a filtrar|Sheet1|
|Resultado detalhado|Verifique para obter o resultado detalhado|True|
|Atribuir resultado à variável||Variável|

### Excluir linha/coluna
  
Comando para deletar linhas e/ou colunas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha |Nome da folha da qual a linha ou coluna será excluída|Sheet1|
|Linha(s)|Intervalo de linhas a serem excluídas|1:5|
|Coluna(s)|Intervalo de colunas a serem removidas|A:G|

### Inserir Linha/Coluna
  
Comando para inserir linhas e/ou colunas
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha |Nome da folha da qual a linha ou coluna será excluída|Sheet1|
|Linha(s)|Intervalo de linhas a serem excluídas|1:5|
|Coluna(s)|Intervalo de colunas a serem removidas|A:G|
