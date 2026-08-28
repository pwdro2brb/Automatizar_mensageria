# Uber 3: Criar Rascunhos de E-mail

## Objetivo

Este robô cria um rascunho de e-mail para cada responsável por centro de custo, permitindo que cada responsável analise os valores e as utilizações correspondentes.

Os rascunhos são criados automaticamente no Outlook, com os arquivos relacionados a cada responsável anexados ao respectivo e-mail.

## Pré-requisitos

Antes de executar o Uber 3, conclua estes processos na seguinte ordem:

1. Uber 1: Atualizar Responsáveis.
2. Uber 2: Gerar Relatórios e Pastas.
3. Uber 3: Criar Rascunhos de E-mail.

O Uber 3 depende diretamente dos arquivos produzidos pelo Uber 2.

## Localização dos arquivos

O robô procura os arquivos na pasta:

AUTOMATIZAR_MENSAGERIA/dist/arquivos/uber/

Dentro desse caminho, o robô identifica automaticamente a pasta mais recente gerada pelo Uber 2.

Essa pasta deve conter os documentos separados por responsável.

## Funcionamento

Durante a execução, o robô:

1. Localiza a pasta mais recente do processo Uber.
2. Identifica os arquivos de cada responsável.
3. Localiza os destinatários configurados para cada centro de custo.
4. Cria um rascunho de e-mail para cada responsável.
5. Adiciona os arquivos correspondentes como anexos.
6. Insere a mensagem padrão do processo.

## Mensagem padrão

A mensagem utilizada pelo robô segue o padrão definido no código do processo:

[MENSAGEM PADRÃO DO E-MAIL]

## Resultado esperado

Ao finalizar, os e-mails estarão disponíveis na pasta de Rascunhos do Outlook.

O robô não envia os e-mails automaticamente. Antes do envio, revise:

- o destinatário;
- o assunto;
- a mensagem;
- os anexos;
- os centros de custo relacionados.

## Possíveis problemas

### Nenhuma pasta foi encontrada

Verifique se o Uber 2 foi executado e se uma nova pasta foi criada dentro de:

AUTOMATIZAR_MENSAGERIA/dist/arquivos/uber/

### Alguns responsáveis não receberam rascunho

Verifique se:

- o responsável está preenchido na planilha;
- o e-mail está cadastrado corretamente;
- o Uber 1 localizou o responsável;
- os arquivos foram gerados corretamente pelo Uber 2.

### O Outlook não criou os rascunhos

Confirme se o Outlook está:

- instalado;
- aberto ou disponível para automação;
- conectado à conta corporativa;
- sem janelas de confirmação bloqueando o processo.