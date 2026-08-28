# Gerar Relatório de Envio para os Correios

## Objetivo

Este robô consulta no Agilis as correspondências enviadas durante o período selecionado e gera uma planilha editada no padrão utilizado pelos Correios.

O relatório pode ser gerado considerando os envios do período da manhã ou do período da tarde.

## Pré-requisitos

Para executar o robô, é necessário:

- possuir uma chave válida da API do Agilis;
- configurar a chave da API no Hub;
- possuir conexão com a internet;
- possuir acesso às informações do Agilis;
- manter o Excel disponível para geração do resultado.

## Configuração da API do Agilis

Este processo utiliza a API do Agilis para consultar os registros de envio.

A chave precisa estar preenchida na aba Configurações do Hub.

A aba Ajuda possui uma seção chamada `Chave API Agilis`, com as instruções para gerar e configurar essa chave.

## Períodos processados

O relatório considera as correspondências enviadas em um dos períodos utilizados pelo processo:

- manhã;
- tarde.

Antes de executar, confira qual período deve ser processado para evitar duplicidade ou ausência de registros.

## Funcionamento

Durante a execução, o robô:

1. Conecta-se à API do Agilis.
2. Consulta os registros de correspondências enviadas.
3. Considera o período da manhã ou da tarde.
4. Obtém as informações necessárias para o relatório.
5. Organiza os registros encontrados.
6. Ajusta as colunas e os dados para o padrão dos Correios.
7. Gera a planilha final.

## Resultado esperado

Ao final, será criada uma planilha contendo as correspondências enviadas no período processado, formatada de acordo com o padrão utilizado pelos Correios.

Antes de encaminhar ou utilizar o arquivo, confira:

- a data considerada;
- o período processado;
- a quantidade de registros;
- os nomes dos destinatários;
- os endereços;
- os códigos ou serviços informados;
- a formatação final da planilha.

## Cuidados importantes

- Confirme o período antes da execução.
- Evite executar duas vezes para o mesmo período sem revisar o resultado.
- Não compartilhe a chave da API.
- Verifique se os registros no Agilis foram preenchidos corretamente.
- Revise a planilha antes de enviá-la aos Correios.

## Possíveis problemas

### A chave da API do Agilis não está configurada

Acesse a aba Configurações e preencha o campo `Chave API Agilis`.

Consulte a seção correspondente na ajuda geral do Hub.

### Nenhum registro foi encontrado

Verifique se:

- existem correspondências enviadas no período;
- a data consultada está correta;
- os registros foram preenchidos no Agilis;
- a API do Agilis está acessível;
- a chave configurada continua válida.

### A planilha foi gerada sem algumas informações

Confira se os campos obrigatórios foram preenchidos corretamente nos chamados ou registros do Agilis.

### Registros duplicados

Verifique se o mesmo período foi processado mais de uma vez e compare os arquivos gerados antes de utilizar o resultado.