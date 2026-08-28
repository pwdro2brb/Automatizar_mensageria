# Cobrança de Boletos

## Objetivo

Este robô analisa a caixa de e-mail para localizar processos de cobrança nos quais o boleto ainda não foi enviado em anexo.

Quando identifica uma loja que está há cinco dias sem anexar o boleto na sequência de e-mails, o robô cria um rascunho de cobrança no Outlook.

## Pré-requisitos

Não é necessário fornecer uma planilha ou arquivo de entrada para executar este robô.

Entretanto, é necessário que:

- o Outlook esteja instalado;
- a conta corporativa esteja conectada;
- a caixa de e-mail esteja acessível;
- os e-mails do processo estejam disponíveis;
- a automação tenha permissão para acessar o Outlook.

## Funcionamento

Durante a execução, o robô:

1. Acessa a caixa de e-mail.
2. Localiza os e-mails relacionados às cobranças.
3. Analisa as conversas dos últimos dias.
4. Verifica a sequência de respostas de cada processo.
5. Verifica se um boleto foi enviado em anexo.
6. Calcula o tempo sem retorno.
7. Identifica as lojas que estão há cinco dias sem enviar o boleto.
8. Cria um rascunho de cobrança com a mensagem padrão.

## Regra de cobrança

O robô cria uma nova cobrança quando:

- o processo possui um e-mail de cobrança anterior;
- ainda não existe um boleto anexado na sequência analisada;
- passaram-se cinco dias sem o envio do boleto.

## Mensagem padrão

O texto utilizado no rascunho segue o modelo configurado no código:

[MENSAGEM PADRÃO DE COBRANÇA]

## Resultado esperado

Os e-mails de cobrança serão criados na pasta de Rascunhos do Outlook.

O robô não envia os e-mails automaticamente.

Antes do envio, confira:

- o destinatário;
- a loja relacionada;
- o histórico da conversa;
- a quantidade de dias sem retorno;
- a existência de anexos;
- o texto da cobrança.

## Possíveis problemas

### Nenhum rascunho foi criado

Isso pode significar que:

- não existem cobranças pendentes;
- os boletos já foram anexados;
- ainda não passaram cinco dias;
- os e-mails não foram identificados pelo assunto esperado;
- a caixa de e-mail não estava acessível.

### O Outlook apresentou erro

Feche janelas de confirmação, verifique a conexão da conta e execute novamente.

### Um boleto existente não foi identificado

Confira se o boleto foi enviado como anexo em uma mensagem pertencente à mesma sequência de e-mails analisada pelo robô.