# Incluir Correspondências Rápidas

## Objetivo

Este robô registra automaticamente no Podio as correspondências recebidas pelo setor.

Como as informações físicas das correspondências não podem ser identificadas automaticamente pelo computador, os dados precisam ser preenchidos previamente na planilha `encomendas.xlsx`.

Depois do preenchimento, o robô lê cada linha da planilha e inclui as correspondências no Podio de forma automática e mais rápida.

## Pré-requisitos

Para executar o robô, é necessário:

- possuir acesso ao Podio;
- configurar as chaves da API do Podio;
- preencher a planilha `encomendas.xlsx`;
- fechar a planilha antes de iniciar o processo;
- possuir conexão com a internet.

## Configuração da API do Podio

Este robô depende da API do Podio para registrar as correspondências.

As seguintes informações precisam estar configuradas no Hub:

- Podio Client ID;
- Podio Client Secret;
- Podio App ID;
- Podio App Token.

A aba Ajuda possui uma seção chamada `Chave API Podio`, com as instruções para obtenção e configuração dessas informações.

## Localização da planilha

A planilha utilizada pelo processo está localizada em:

arquivos/encomendas/encomendas.xlsx

## Preparação da planilha

Antes de incluir novas correspondências:

1. Abra a planilha `encomendas.xlsx`.
2. Mantenha o cabeçalho original.
3. Apague todas as informações antigas abaixo do cabeçalho.
4. Preencha uma linha para cada nova correspondência recebida.
5. Confira as informações preenchidas.
6. Salve a planilha.
7. Feche completamente o Excel antes de executar o robô.

Não exclua nem altere os nomes das colunas do cabeçalho.

## Funcionamento

Durante a execução, o robô:

1. Localiza a planilha `encomendas.xlsx`.
2. Lê as linhas preenchidas abaixo do cabeçalho.
3. Valida as informações disponíveis.
4. Conecta-se ao Podio por meio da API.
5. Inclui cada correspondência no aplicativo correspondente.
6. Apresenta o andamento do processo no console do Hub.

## Resultado esperado

Ao final da execução, as correspondências preenchidas na planilha estarão registradas no Podio.

Após o processamento, confira no Podio:

- a quantidade de correspondências incluídas;
- os destinatários;
- os tipos de correspondência;
- as datas preenchidas;
- os demais dados importados da planilha.

## Cuidados importantes

- Não execute o robô com a planilha aberta.
- Não altere o cabeçalho da planilha.
- Não deixe informações de processos anteriores abaixo do cabeçalho.
- Revise os dados antes de executar.
- Não compartilhe as chaves da API do Podio.

## Possíveis problemas

### A planilha está aberta

Feche completamente a planilha e execute o processo novamente.

### A planilha não foi encontrada

Confirme se o arquivo está neste caminho:

arquivos/encomendas/encomendas.xlsx

Confira também se o nome permanece exatamente como `encomendas.xlsx`.

### A API do Podio não está configurada

Acesse a aba Configurações e preencha os campos da API do Podio.

Consulte a seção `Chave API Podio` na ajuda geral do Hub para obter as instruções de configuração.

### Algumas correspondências não foram incluídas

Verifique se:

- todas as informações obrigatórias foram preenchidas;
- não existem linhas incompletas;
- o cabeçalho original foi mantido;
- as chaves da API estão corretas;
- o Podio está acessível;
- a estrutura do aplicativo no Podio não foi alterada.