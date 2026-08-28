# Faturamento dos Correios

Este arquivo documenta os dois processos de faturamento disponíveis no Hub:

1. Faturamento 1: Gerar Rascunhos.
2. Faturamento 2: Processo Completo.

---

# Faturamento 1: Gerar Rascunhos

## Objetivo

Este robô acessa a pasta pública de faturamento e cria um rascunho de e-mail para cada regional, exceto Belo Horizonte.

Cada rascunho contém uma mensagem padrão e os arquivos encontrados na pasta da respectiva regional.

O robô identifica automaticamente os destinatários de acordo com a regional processada.

## Localização dos arquivos

O robô utiliza a seguinte estrutura da pasta pública:

Correios/Faturamento/ANO ATUAL/MÊS MAIS RECENTE/REGIONAL/

Exemplo:

Correios/Faturamento/2026/Agosto/Rio de Janeiro/

A regional de Belo Horizonte não participa dessa etapa.

## Arquivos necessários

Cada pasta regional deve conter:

1. Um boleto em formato PDF.
2. Uma planilha dos Correios em formato Excel.

A planilha dos Correios normalmente possui nome numérico, por exemplo:

1234567.xlsx

## Funcionamento

Durante a execução, o robô:

1. Localiza o ano atual.
2. Localiza a pasta mais recente do faturamento.
3. Percorre as pastas das regionais.
4. Ignora a pasta de Belo Horizonte.
5. Identifica o boleto em PDF.
6. Identifica a planilha dos Correios.
7. Determina os destinatários da regional.
8. Cria um rascunho no Outlook.
9. Adiciona a mensagem padrão.
10. Anexa os arquivos da regional.

## Mensagem padrão

O rascunho utiliza o modelo definido no processo:

[MENSAGEM PADRÃO DO FATURAMENTO 1]

## Resultado esperado

Será criado um rascunho para cada regional que possuir os arquivos necessários.

Antes do envio, confira:

- os destinatários;
- a regional;
- o assunto;
- a mensagem;
- o boleto;
- a planilha dos Correios.

## Possíveis problemas

### Uma regional não gerou rascunho

Verifique se a pasta regional contém:

- um arquivo PDF;
- uma planilha Excel dos Correios;
- os nomes e formatos esperados.

### O destinatário não foi localizado

Confira se a regional está contemplada no mapeamento de destinatários utilizado pelo código.

---

# Faturamento 2: Processo Completo

## Objetivo

Este robô analisa as respostas recebidas por e-mail e prepara os arquivos necessários para o lançamento da nota fiscal.

O processo utiliza:

- o arquivo Excel recebido por e-mail;
- o arquivo dos Correios presente na pasta da regional;
- o boleto;
- as informações da pasta pública.

Após preparar o Rateio Pag, o robô inicia o lançamento da nota fiscal automaticamente.

A ação final de confirmação não é executada, permitindo que o usuário revise os dados antes de concluir o lançamento.

## E-mails analisados

O robô procura respostas com o assunto:

RES: Extrato Correios

A sequência de e-mails deve conter o arquivo Excel necessário para a continuidade do faturamento.

## Funcionamento

Durante a execução, o robô:

1. Acessa a caixa de entrada do Outlook.
2. Localiza os e-mails com o assunto esperado.
3. Analisa as respostas recebidas.
4. Verifica se existe um arquivo Excel anexado.
5. Identifica a regional relacionada ao e-mail.
6. Localiza a pasta correspondente na pasta pública.
7. Localiza o arquivo dos Correios.
8. Combina o arquivo recebido com o arquivo da regional.
9. Gera o Rateio Pag.
10. Utiliza o Rateio Pag e o boleto no processo de lançamento da nota fiscal.
11. Preenche os dados necessários no sistema.
12. Interrompe antes da confirmação final.

## O que é o Rateio Pag

O Rateio Pag é o arquivo utilizado como padrão de rateio no sistema de lançamento de notas fiscais.

O arquivo relaciona os valores da cobrança aos respectivos centros de custo e permite que o lançamento seja preparado corretamente.

## Resultado esperado

Ao final do processo:

- o Rateio Pag estará criado;
- os arquivos da regional estarão organizados;
- o lançamento da nota fiscal estará preenchido;
- a confirmação final ficará pendente para conferência manual.

## Conferência obrigatória

Antes de concluir o lançamento, confira:

- o fornecedor;
- a regional;
- o número do documento;
- o valor total;
- o vencimento;
- o boleto;
- o Rateio Pag;
- os centros de custo;
- os valores distribuídos;
- os anexos inseridos.

## Possíveis problemas

### E-mail não identificado

Confirme se o assunto contém o padrão esperado:

RES: Extrato Correios

### Arquivo Excel não encontrado

Verifique se a resposta possui o arquivo em anexo e se o formato é `.xlsx`.

### Regional não identificada

Confira se o assunto, o remetente ou os anexos possuem as informações utilizadas pelo robô para identificar a regional.

### Rateio Pag não foi gerado

Verifique a estrutura dos arquivos, os centros de custo e os valores informados nas planilhas de entrada.