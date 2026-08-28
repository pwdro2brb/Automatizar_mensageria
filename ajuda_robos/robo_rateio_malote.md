# Rateio de Malote

## Objetivo

Este robô realiza o rateio dos custos de utilização do serviço de malote entre os respectivos centros de custo.

O processo identifica quais centros de custo utilizaram o serviço e distribui os valores de acordo com as informações encontradas nos arquivos dos Correios e do Agilis.

Esse processo costuma ser complexo e demorado quando realizado manualmente. A automação reduz o tempo de execução e ajuda a padronizar a distribuição dos custos.

## Pré-requisitos

Para executar o robô, é necessário:

- possuir acesso ao site Malote Web dos Correios;
- configurar a senha do Malote Web na área de Configurações do Hub;
- possuir o arquivo de cobrança dos Correios;
- possuir o relatório extraído do Agilis;
- manter os arquivos na pasta pública correta.

## Localização dos arquivos

Os arquivos devem estar na pasta pública, seguindo aproximadamente esta estrutura:

Correios/Faturamento/ANO ATUAL/MÊS MAIS RECENTE/BH/

Exemplo:

Correios/Faturamento/2026/Agosto/BH/

O nome exato das pastas pode variar de acordo com a organização adotada no diretório público.

## Arquivos necessários

A pasta de Belo Horizonte deve conter:

1. Arquivo dos Correios em formato Excel.
2. Relatório do Agilis.

O arquivo dos Correios normalmente possui um nome numérico, por exemplo:

1234567.xlsx

## Funcionamento

Durante a execução, o robô:

1. Localiza a pasta do ano atual.
2. Localiza a pasta mais recente do faturamento.
3. Entra na pasta de Belo Horizonte.
4. Identifica o arquivo dos Correios.
5. Identifica o relatório do Agilis.
6. Consulta as informações necessárias do Malote Web.
7. Relaciona as utilizações aos respectivos centros de custo.
8. Realiza a distribuição dos valores.
9. Gera o arquivo final de rateio na pasta pública.

## Resultado esperado

Ao final do processo, será criado um arquivo de rateio de malote na pasta pública correspondente ao período processado.

Antes de utilizar o resultado, confira:

- o período considerado;
- o valor total do documento;
- os centros de custo;
- os valores distribuídos;
- os registros que não foram localizados.

## Possíveis problemas

### Credencial do Malote Web não configurada

Acesse a seção Configurações do Hub e preencha o campo correspondente à senha do Malote Web.

### Arquivo dos Correios não encontrado

Verifique se:

- o arquivo está na pasta de Belo Horizonte;
- o arquivo está em formato `.xlsx`;
- o nome do arquivo é predominantemente numérico;
- o arquivo não está aberto no Excel.

### Relatório do Agilis não encontrado

Confirme se o relatório foi salvo na mesma pasta utilizada pelo processo e se o arquivo está no formato esperado.

### Centro de custo não identificado

Revise os registros não localizados e confirme se as informações de referência estão preenchidas corretamente no Agilis.