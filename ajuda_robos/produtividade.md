# Gerar Produtividade

## Objetivo

Este robô realiza a distribuição da produtividade de cada colaborador do setor.

Para produzir o resultado, a automação consolida informações de quatro fontes referentes ao mês anterior:

- Agilis;
- SAP ou Bússola;
- Lançamentos;
- Podio.

Além dos quatro relatórios, o robô utiliza a planilha-base de produtividade para organizar e preencher os resultados.

## Modos de execução

O robô possui dois modos de execução.

### Processo completo

Nesse modo, o robô:

1. Acessa os sistemas necessários.
2. Baixa os relatórios.
3. Localiza os arquivos na pasta Downloads.
4. Move ou utiliza os arquivos baixados.
5. Realiza as edições e consolidações.
6. Atualiza a planilha de produtividade.

Durante as etapas que utilizam SAP, Bússola ou automação de tela, não mexa no mouse ou no teclado.

Algumas páginas podem solicitar autenticação MFA.

### Apenas edição

Utilize essa opção quando os relatórios já tiverem sido baixados.

Os arquivos devem estar na pasta:

AUTOMATIZAR_MENSAGERIA/dist/arquivos/produtividade/

Não é necessário renomear os relatórios. O robô tentará identificar os arquivos de acordo com seu conteúdo e formato.

## Arquivos necessários

Para executar apenas a edição, coloque na pasta de produtividade:

1. Relatório do Agilis.
2. Relatório do SAP ou Bússola.
3. Planilha de lançamentos.
4. Relatório do Podio.
5. Planilha-base de produtividade.

Todos os relatórios devem ser referentes ao mês anterior ao mês de execução.

## Funcionamento

Durante o processamento, o robô:

1. Identifica os arquivos disponíveis.
2. Confere o período dos relatórios.
3. Lê os dados do Agilis.
4. Lê os dados do SAP ou Bússola.
5. Lê a planilha de lançamentos.
6. Lê os dados do Podio.
7. Relaciona os registros aos colaboradores.
8. Consolida os indicadores.
9. Preenche a planilha-base de produtividade.
10. Gera ou atualiza o arquivo final.

## Resultado esperado

O resultado será uma planilha de produtividade com os dados consolidados por colaborador.

Após o processamento, revise:

- os nomes dos colaboradores;
- os dados do Podio;
- os dados do Agilis;
- os dados do SAP ou Bússola;
- os lançamentos;
- os registros que não foram relacionados;
- os valores totais de cada indicador.

## Possíveis problemas

### Um relatório não foi identificado

Verifique se:

- o arquivo está na pasta correta;
- o arquivo está em formato compatível;
- o relatório corresponde ao mês anterior;
- o arquivo não está corrompido;
- a estrutura original do relatório não foi alterada.

### O robô não localizou a planilha-base

Confirme se existe um arquivo de produtividade no mesmo caminho dos demais relatórios.

### O download automático falhou

Você pode baixar os quatro relatórios manualmente, colocá-los na pasta de produtividade e executar a opção Apenas edição.

### O computador parece estar sendo controlado

Algumas extrações utilizam automação de interface. Não mexa no mouse ou no teclado enquanto essas etapas estiverem em andamento.