# Uber 1 e Uber 2

Este arquivo documenta os dois primeiros processos do fluxo Uber:

1. Uber 1: Atualizar Responsáveis.
2. Uber 2: Gerar Relatórios e Pastas.

A ordem recomendada de execução é:

1. Uber 1.
2. Conferência da planilha de responsáveis.
3. Uber 2.
4. Uber 3.

---

# Uber 1: Atualizar Responsáveis

## Objetivo

Este processo atualiza a planilha de responsáveis por centro de custo.

O robô combina informações do SAP, da base de ativos e da planilha atual de responsáveis para localizar o responsável correspondente a cada centro de custo.

## Localização dos arquivos

Todos os arquivos devem estar na pasta:

AUTOMATIZAR_MENSAGERIA/dist/arquivos/uber/

## Arquivos necessários

A pasta deve conter:

1. Planilha extraída da transação KS13 do SAP, com todos os centros de custo.
2. Arquivo mais recente da base de ativos.
3. Planilha de responsáveis por centro de custo.

## Funcionamento

Durante a execução, o robô:

1. Localiza a planilha da transação KS13.
2. Localiza o arquivo mais recente da base de ativos.
3. Localiza a planilha de responsáveis por centro de custo.
4. Relaciona os centros de custo às informações disponíveis.
5. Localiza os responsáveis correspondentes.
6. Atualiza a planilha de responsáveis.
7. Gera o arquivo atualizado para utilização no Uber 2.

## Conferência recomendada

Após o processamento, revise a planilha gerada.

Alguns responsáveis podem não ser encontrados automaticamente se:

- o nome estiver escrito de forma diferente;
- o cadastro não existir na base de ativos;
- o centro de custo estiver ausente;
- existirem dados duplicados;
- a estrutura da planilha tiver sido alterada.

Corrija manualmente os registros não localizados antes de executar o Uber 2.

## Resultado esperado

O robô criará ou atualizará a planilha de responsáveis por centro de custo.

---

# Uber 2: Gerar Relatórios e Pastas

## Objetivo

Este processo organiza a utilização mensal do Uber por responsável e por centro de custo.

O robô utiliza a planilha produzida pelo Uber 1 e o relatório do Uber referente ao mês anterior.

## Pré-requisitos

Antes de executar o Uber 2:

1. Execute o Uber 1.
2. Revise a planilha atualizada de responsáveis.
3. Corrija os responsáveis não localizados.
4. Coloque o relatório do Uber na pasta indicada.

## Arquivos necessários

A pasta deve conter:

1. Planilha atualizada pelo Uber 1.
2. Relatório do Uber referente ao mês anterior.

Os arquivos devem estar em:

AUTOMATIZAR_MENSAGERIA/dist/arquivos/uber/

## Funcionamento

Durante a execução, o robô:

1. Localiza a planilha atualizada de responsáveis.
2. Localiza o relatório mensal do Uber.
3. Relaciona as utilizações aos centros de custo.
4. Relaciona os centros de custo aos responsáveis.
5. Consolida os dados para envio.
6. Cria o arquivo consolidado.
7. Cria uma pasta para o período processado.
8. Separa os arquivos por responsável.

## Principal arquivo gerado

O principal arquivo produzido pelo processo é:

consolidado_para_envio_ATUALIZADO

Esse arquivo contém os dados consolidados que serão utilizados nas etapas seguintes.

## Pastas geradas

O robô também cria uma pasta contendo a utilização de cada centro de custo separada por responsável.

Essa estrutura é utilizada pelo Uber 3 para criar os rascunhos de e-mail.

## Resultado esperado

Ao finalizar, estarão disponíveis:

- o arquivo `consolidado_para_envio_ATUALIZADO`;
- a pasta do período processado;
- os arquivos separados por responsável;
- os documentos necessários para o Uber 3.

## Possíveis problemas

### O responsável não foi identificado

Retorne ao resultado do Uber 1 e confira a planilha de responsáveis.

### O relatório do Uber não foi encontrado

Confirme se o arquivo está na pasta correta e corresponde ao mês anterior.

### O Uber 3 não encontrou os arquivos

Verifique se o Uber 2 criou a pasta do período e se existem arquivos separados por responsável.