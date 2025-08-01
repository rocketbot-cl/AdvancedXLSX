



# Opções avançadas para XLSX
  
Formate células, crie e remova planilhas, filtre dados, adicione e exclua colunas e linhas, abra arquivos xls e transforme-os no formato xlsx.  

*Read this in other languages: [English](Manual_AdvancedXLSX.md), [Português](Manual_AdvancedXLSX.pr.md), [Español](Manual_AdvancedXLSX.es.md)*
  
![banner](imgs/Banner_AdvancedXLSX.png)
## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  



## Como usar este módulo

Apenas se você estiver utilizando a versão 2023 do Rocketbot, siga os passos abaixo para evitar o erro:

ImportError: cannot import name 'etree' from 'lxml'

1. Dirija-se à pasta raiz do Rocketbot e verifique se a biblioteca 'lxml' existe.
2. Caso não exista, a partir de um terminal, vá para a pasta raiz do Rocketbot e digite:  pip install lxml -t .
3. Leve em consideração que você deve instalar a biblioteca com o Python 3.10 de 64 bits.

## Descrição do comando

### Abrir xls
  
Abra um arquivo xls para trabalhar com o comando nativo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo XLS|Selecione o arquivo XLS que deseja abrir|example.xls|
|Coluna/as como data (opcional)||0|
|Id (optional) |Identificador de sessão|id|
|Encoding|Tipo de Encoding a aplicar. Por padrão latin-1|latin-1|
|Atribuir resultado à variável||Variável|

### Abrir xlsx avançado
  
Abra um arquivo xlsx para trabalhar com o comando nativo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo XLSX|Selecione o arquivo XLSX que deseja abrir|example.xlsx|
|Somente leitura|Marque se deseja abrir o xlsx somente para leitura, o conteúdo não pode ser editado.|False|
|Manter vba|Marque para manter o possível código VBA que poderia estar no workbook.|False|
|Somente dados|Controla se as celulas com fórmulas tiverem a fórmula (definido) ou o valor armazenado a|False|
|Manter links|Marcar se devem manter os enlaces aos libros de trabalho externos.|False|
|Id (optional) |Identificador de sessão|id|
|Atribuir resultado à variável||Variável|

### Converter xls para xlsx
  
Converter um arquivo de formato xls para o formato xlsx
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo XLS|Selecione o arquivo XLS que deseja abrir|path/to/file/example.xls|
|Caminho do arquivo XLSX|Coloque o caminho completo onde deseja salvar o arquivo XLSX (incluindo nome e extensão '.xlsx')|path/to/file/example.xlsx|
|Encoding|Tipo de Encoding a aplicar. Por padrão latin-1|latin-1|

### Converter planilha em csv
  
Converta uma planilha do arquivo xlsx aberto em csv
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho do arquivo CSV|Selecione o arquivo CSV que deseja abrir|path/to/file/example.csv|
|Delimitador|Delimitador da arquivo csv|,|
|Formato de saída de data|Formato com o qual as datas da planilha xlsx serão convertidas para csv|%d/%m/%Y|
|Atribuir resultado a variável |Nome da variável para armazenar o resultado|Variável|

### Escrever em célula
  
Escreve um valor em uma célula específica. Se for passada uma matriz, escreve cada valor verticalmente na mesma coluna.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da planilha||Planilha1|
|Célula||A1|
|Valor a escrever|Pode ser um único valor ou uma matriz vertical (ex [[1],[0],[1]]).|42 ou [[1],[0],[1]]|

### Ler intervalo
  
Retorna o valor do intervalo fornecido. Um valor se o intervalo for uma célula ou uma lista se o intervalo tiver diversas células.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha |Nome da planilha onde o intervalo está localizado|Sheet1|
|Célula inicial|Célula ou intervalo|A1|
|Atribuir resultado à variável (Coluna)|Nome da variável onde o comprimento da coluna será salvo|Variável|

### Renomear folha
  
Renomear uma folha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha a ser renomeada |Nome da folha a ser renomeada|OldSheet|
|Novo nome da folha|Nome da folha|NewSheet|

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

### Proteger planilha
  
Proteger planilha
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome da folha |Nome da folha|Sheet2|
|Senha|Senha que será aplicada à planilha|1524|
|Atribuir resultado à variável|Nome da variável onde será salvo|Variável|

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
|Filtros|Filtros a aplicar. Para filtrar por vazio usar == None|["A > 3", "D *ARS", "C == Invoice"]|
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

### Inserir imagem
  
Inserir uma imagem em um documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho da imagem|Selecione o arquivo de imagem que deseja inserir no documento|example.png|
|Folha |Nome da folha de documento onde inserir a imagem|Sheet1|
|Célula |Célula onde inserir a imagem|A1|

### Fechar xlsx
  
Fechar um arquivo xlsx aberto
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Atribuir resultado à variável||Variável|
