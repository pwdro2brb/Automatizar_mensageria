# Atualizar Responsáveis de Viagens

## Objetivo

Este robô atualiza os responsáveis cadastrados na planilha de centros de custo de viagens.

O processo consulta o responsável de cada centro de custo no arquivo EXPORT do SAP. Em seguida, utiliza a Base de Ativos para identificar o nome completo do responsável e localizar o respectivo superior imediato.

## Pasta do processo

Os arquivos devem ser colocados na pasta:

`uber\Responsaveis Viagens`

Essa pasta fica dentro da pasta principal de arquivos configurada no Hub.

Ao clicar em **Cancelar** no lembrete de execução, o Hub abrirá automaticamente essa pasta.

## Arquivos necessários

Antes de executar o robô, coloque os três arquivos abaixo na pasta do processo.

### 1. Arquivo EXPORT do SAP

O nome do arquivo deve conter a palavra:

`export`

Exemplo:

`EXPORT_20260901_101748.xlsx`

O arquivo deve possuir as colunas referentes ao centro de custo e ao responsável.

### 2. Base de Ativos

O nome do arquivo deve conter:

`Base de Ativos`

Exemplo:

`Base de Ativos 01.09.2026.xlsx`

A Base de Ativos deve possuir, entre outras, as seguintes colunas:

- Nome Funcionário
- Nome Superior
- Nome da Função
- E-mail

### 3. Planilha de Centro de Custo Viagens

O nome do arquivo deve conter:

`centro de custo viagens`

Exemplo:

`centro de custo viagens.xlsx`

Esta é a planilha que será atualizada pelo robô.

## Antes de executar

1. Atualize os três arquivos na pasta do processo.
2. Feche todas as planilhas no Excel.
3. Confirme se o EXPORT utilizado é o mais recente.
4. Confirme se a Base de Ativos utilizada é a mais recente.
5. Clique em **Executar** no robô Responsáveis Viagens.
6. No lembrete apresentado pelo Hub, clique em **OK**.

Se precisar conferir ou substituir os arquivos, clique em **Cancelar**. O Hub abrirá a pasta do processo.

## Funcionamento

O robô realiza as seguintes etapas:

1. Localiza automaticamente o arquivo EXPORT mais recente.
2. Localiza automaticamente a Base de Ativos mais recente.
3. Localiza a planilha de centro de custo viagens.
4. Lê o código da coluna **Centro de Custo sem Divisão**.
5. Procura o centro de custo no EXPORT.
6. Obtém o responsável cadastrado no SAP.
7. Procura esse responsável na Base de Ativos.
8. Obtém o nome completo do responsável.
9. Obtém o Nome Superior na mesma linha do funcionário.
10. Atualiza o 1º aprovador.
11. Atualiza o superior do 1º aprovador.
12. Adiciona uma coluna de status ao final da planilha.
13. Salva um novo arquivo atualizado.

## Colunas atualizadas

### Coluna E

A coluna E corresponde ao:

`1º Aprovador`

Ela recebe o nome completo do responsável localizado na Base de Ativos.

### Coluna F

A coluna F corresponde ao:

`Superior do 1º Aprovador`

Ela recebe o conteúdo da coluna **Nome Superior** encontrada na mesma linha do responsável na Base de Ativos.

### Última coluna

O robô cria ou atualiza a coluna:

`Status Atualização SAP`

Essa coluna informa o resultado do processamento de cada centro de custo.

## Status possíveis

### OK

O 1º aprovador e o superior do 1º aprovador já estavam corretos.

### Alterado: 1º aprovador atualizado

Somente o conteúdo da coluna E foi atualizado.

### Alterado: superior atualizado

Somente o conteúdo da coluna F foi atualizado.

### Alterado: 1º aprovador atualizado | superior atualizado

As colunas E e F foram atualizadas.

### Centro de custo não encontrado no SAP

O código da coluna **Centro de Custo sem Divisão** não foi localizado no arquivo EXPORT.

Esse registro deve ser revisado manualmente.

### Nome não encontrado na Base de Ativos

O centro de custo foi localizado no EXPORT, mas o responsável informado pelo SAP não foi localizado com segurança na Base de Ativos.

Esse registro deve ser revisado manualmente.

### Nome Superior vazio na Base de Ativos

O responsável foi localizado na Base de Ativos, mas o campo **Nome Superior** está vazio.

Nesse caso, o preenchimento da coluna F deve ser revisado manualmente.

## Arquivo gerado

Ao final do processo, será criado o arquivo:

`Centro_Custo_Viagens_Atualizado.xlsx`

O arquivo será salvo na pasta:

`uber\Responsaveis Viagens`

A planilha original não será substituída.

## Conferência do resultado

Após a execução:

1. Abra o arquivo `Centro_Custo_Viagens_Atualizado.xlsx`.
2. Localize a coluna `Status Atualização SAP`.
3. Ative ou utilize o filtro da planilha.
4. Confira os registros com status `Alterado`.
5. Revise os casos de centro de custo não encontrado no SAP.
6. Revise os nomes não encontrados na Base de Ativos.
7. Revise os casos em que o Nome Superior está vazio.
8. Depois da conferência, salve ou encaminhe o arquivo conforme o procedimento da área.

## Cuidados

- Mantenha as planilhas fechadas durante a execução.
- Não remova as colunas principais dos arquivos.
- Não altere os cabeçalhos usados pelo processo.
- Não renomeie o arquivo de saída para um nome que contenha exatamente o padrão do arquivo de entrada.
- Se houver mais de um EXPORT, o robô utilizará o arquivo mais recente.
- Se houver mais de uma Base de Ativos, o robô utilizará o arquivo mais recente.
- Confira sempre se os arquivos mais recentes são realmente os arquivos corretos.

## Em caso de erro

Verifique:

1. Se a pasta `Responsaveis Viagens` existe dentro da pasta `uber`.
2. Se o arquivo EXPORT contém a palavra `export` no nome.
3. Se a Base de Ativos contém `Base de Ativos` no nome.
4. Se a planilha de viagens contém `centro de custo viagens` no nome.
5. Se todas as planilhas estão fechadas no Excel.
6. Se os arquivos não estão corrompidos.
7. Se as colunas obrigatórias ainda existem.
8. Se o usuário possui permissão para ler e gravar arquivos na pasta.

Depois de corrigir os arquivos, execute novamente o robô pelo Hub.