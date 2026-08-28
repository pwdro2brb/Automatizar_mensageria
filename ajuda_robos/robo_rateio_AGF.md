# Rateio AGF

## Objetivo

Este robô realiza o rateio dos serviços dos Correios, distribuindo os valores cobrados entre os respectivos centros de custo.

A automação combina as informações do arquivo dos Correios com os dados obtidos por meio do relatório e da API do Agilis.

## Pré-requisitos

Para executar o robô, é necessário:

- possuir uma chave válida da API do Agilis;
- configurar a chave da API no Hub;
- possuir acesso à pasta pública de faturamento;
- manter o arquivo dos Correios na pasta de Belo Horizonte;
- manter o relatório do Agilis na mesma pasta;
- fechar os arquivos no Excel antes da execução.

## Configuração da API do Agilis

A chave da API do Agilis precisa estar configurada na aba Configurações.

A aba Ajuda possui uma seção chamada `Chave API Agilis`, com o passo a passo para gerar, copiar e salvar a chave.

Não compartilhe a chave da API com outras pessoas.

## Localização dos arquivos

Os arquivos devem estar na pasta de Belo Horizonte correspondente ao período mais recente:

Correios/Faturamento/ANO ATUAL/MÊS MAIS RECENTE/BH/

Exemplo:

Correios/Faturamento/2026/Agosto/BH/

## Arquivos necessários

A pasta de Belo Horizonte deve conter:

1. Arquivo dos Correios em formato Excel.
2. Relatório do Agilis.

O arquivo dos Correios normalmente possui um nome numérico, por exemplo:

1234567.xlsx

## Funcionamento

Durante a execução, o robô:

1. Localiza a pasta do ano atual.
2. Identifica a pasta mais recente de faturamento.
3. Acessa a pasta de Belo Horizonte.
4. Localiza o arquivo dos Correios.
5. Localiza o relatório do Agilis.
6. Consulta as informações necessárias por meio da API do Agilis.
7. Relaciona os registros aos centros de custo.
8. Distribui os valores entre os centros de custo correspondentes.
9. Gera o resultado do rateio.

## Resultado esperado

Ao final, será criado o arquivo de rateio com os valores distribuídos entre os respectivos centros de custo.

Antes de utilizar o resultado, confira:

- o período processado;
- o valor total do arquivo dos Correios;
- os centros de custo;
- os valores distribuídos;
- os registros não identificados;
- a correspondência entre o Agilis e o arquivo dos Correios.

## Cuidados importantes

- Mantenha os dois arquivos na pasta de Belo Horizonte.
- Não execute o processo com os arquivos abertos no Excel.
- Confirme se os arquivos correspondem ao mesmo período.
- Não altere a estrutura original dos relatórios.
- Verifique se a chave da API permanece válida.

## Possíveis problemas

### A chave da API do Agilis não está configurada

Acesse a aba Configurações e preencha o campo `Chave API Agilis`.

Consulte a seção `Chave API Agilis` na ajuda geral do Hub.

### O arquivo dos Correios não foi encontrado

Verifique se:

- o arquivo está na pasta de Belo Horizonte;
- o arquivo está no formato `.xlsx`;
- o nome do arquivo é predominantemente numérico;
- o arquivo não está aberto no Excel.

### O relatório do Agilis não foi encontrado

Confirme se o relatório está na mesma pasta do arquivo dos Correios e se corresponde ao período processado.

### Alguns centros de custo não foram identificados

Revise os registros sem correspondência e confira se os centros de custo estão preenchidos corretamente nas fontes utilizadas.