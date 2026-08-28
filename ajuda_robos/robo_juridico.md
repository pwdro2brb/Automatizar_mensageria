# Relatório Jurídico Montreal

## Objetivo

Este robô prepara o relatório de correspondências jurídicas recebidas no dia atual.

A automação pode baixar o relatório diretamente do Podio ou utilizar um relatório que já tenha sido baixado.

Após localizar o arquivo, o robô edita a planilha e mantém somente as correspondências jurídicas que precisam ser encaminhadas aos colaboradores responsáveis do setor Jurídico.

## Modos de execução

Ao iniciar o robô, o Hub apresenta duas opções principais.

### Sim: baixar o relatório automaticamente

Nessa opção, o robô:

1. Acessa o Podio.
2. Localiza o relatório de correspondências.
3. Baixa o relatório referente ao dia atual.
4. Localiza o arquivo baixado.
5. Edita a planilha.
6. Mantém somente as correspondências jurídicas.
7. Organiza as informações para os colaboradores responsáveis.

Utilize essa opção quando o relatório ainda não tiver sido baixado.

### Não: utilizar um relatório já baixado

Nessa opção, o robô não acessa o Podio para realizar um novo download.

O robô:

1. Procura o relatório já baixado na pasta Downloads.
2. Identifica o arquivo correspondente.
3. Edita a planilha.
4. Mantém somente as correspondências jurídicas.
5. Organiza as informações para os colaboradores responsáveis.

Utilize essa opção quando o relatório já estiver disponível na pasta Downloads.

## Pré-requisitos

Para executar o processo, é necessário:

- possuir acesso ao Podio;
- possuir conexão com a internet, caso seja utilizado o download automático;
- manter o relatório na pasta Downloads, caso seja utilizada a opção sem download;
- manter o Excel disponível para a edição da planilha;
- atender à solicitação de MFA, caso seja apresentada.

## Arquivo utilizado

O processo utiliza o relatório de correspondências extraído do Podio.

O relatório deve corresponder às encomendas e correspondências recebidas no dia atual.

Na opção de utilizar um arquivo já baixado, mantenha o relatório na pasta Downloads. Não mova o arquivo antes da execução.

## Funcionamento

Durante o processamento, o robô:

1. Obtém ou localiza o relatório do Podio.
2. Abre a planilha para processamento.
3. Analisa os registros presentes no relatório.
4. Identifica as correspondências jurídicas.
5. Remove ou desconsidera os registros que não pertencem ao processo.
6. Organiza os dados de acordo com os colaboradores responsáveis.
7. Gera a planilha final para continuidade do fluxo.

## Resultado esperado

Ao final, será gerada uma planilha contendo somente as correspondências jurídicas consideradas pelo processo.

Antes de utilizar ou encaminhar o relatório, confira:

- a data do relatório;
- os destinatários das correspondências;
- a quantidade de registros;
- os colaboradores responsáveis;
- os itens mantidos após a filtragem;
- os registros que eventualmente não foram classificados.

## Possíveis problemas

### O download automático não foi realizado

Verifique se:

- o Podio está acessível;
- a autenticação foi concluída;
- não existe uma solicitação de MFA pendente;
- o navegador não está bloqueando downloads;
- a conexão com a internet está funcionando.

Se necessário, baixe o relatório manualmente e utilize a opção sem download.

### O relatório já baixado não foi encontrado

Confirme se o arquivo permanece na pasta Downloads e se foi baixado recentemente.

Evite deixar vários relatórios semelhantes na pasta, pois isso pode dificultar a identificação do arquivo correto.

### Nenhuma correspondência jurídica foi encontrada

Confira se:

- o relatório corresponde ao dia atual;
- existem correspondências jurídicas no período;
- as classificações utilizadas no Podio permanecem no padrão esperado;
- a estrutura da planilha não foi alterada.

### A planilha final apresentou registros incorretos

Revise a classificação das correspondências no Podio e confira se os destinatários e responsáveis estão cadastrados corretamente.